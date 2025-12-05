#!/usr/bin/env python3
"""
Quality check script for office2md conversions.

Compares output quality between different converters:
- Default (Mammoth + python-docx)
- Pandoc
- Docling

Usage:
    python scripts/quality_check.py examples/input/doc1.docx
    python scripts/quality_check.py examples/input/doc1.docx --output-dir ./quality_tests
    python scripts/quality_check.py examples/input/doc1.docx --converters default,pandoc
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))


@dataclass
class ImageIssue:
    """Represents an image-related issue."""
    issue_type: str  # 'missing', 'misplaced', 'duplicate', 'metadata_leak'
    location: str    # Where in the document
    details: str     # Additional details


@dataclass
class ConversionResult:
    """Result of a single conversion."""
    converter: str
    success: bool
    time_seconds: float
    output_file: Optional[Path] = None
    images_dir: Optional[Path] = None
    error: Optional[str] = None
    
    # Metrics
    file_size: int = 0
    image_count: int = 0
    heading_count: int = 0
    table_count: int = 0
    list_count: int = 0
    bold_count: int = 0
    link_count: int = 0
    
    # Quality issues
    image_issues: List[ImageIssue] = field(default_factory=list)
    table_issues: List[str] = field(default_factory=list)
    formatting_issues: List[str] = field(default_factory=list)
    
    # Sample content
    headings: List[str] = field(default_factory=list)
    table_sample: str = ""
    
    # Image analysis
    images_in_tables: int = 0
    images_outside_tables: int = 0
    image_placeholders: int = 0  # For docling's <!-- image -->
    image_metadata_leaks: int = 0  # For pandoc's {width="..."}
    consecutive_images: List[Tuple[int, str]] = field(default_factory=list)  # (count, location)


@dataclass
class QualityReport:
    """Complete quality report comparing all converters."""
    input_file: Path
    results: Dict[str, ConversionResult] = field(default_factory=dict)
    
    def add_result(self, result: ConversionResult):
        self.results[result.converter] = result


class QualityChecker:
    """Quality checker for office2md conversions."""
    
    CONVERTERS = {
        "default": [],
        "pandoc": ["--use-pandoc"],
        "docling": ["--use-docling"],
    }
    
    def __init__(
        self,
        input_file: str,
        output_dir: Optional[str] = None,
        converters: Optional[List[str]] = None,
        verbose: bool = False
    ):
        self.input_file = Path(input_file)
        self.output_dir = Path(output_dir) if output_dir else Path("./quality_check_output")
        self.converters = converters or list(self.CONVERTERS.keys())
        self.verbose = verbose
        
        if not self.input_file.exists():
            raise FileNotFoundError(f"Input file not found: {self.input_file}")
        
        # Create output directory
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def check_converter_availability(self) -> Dict[str, bool]:
        """Check which converters are available."""
        availability = {}
        
        # Default is always available
        availability["default"] = True
        
        # Check Pandoc
        availability["pandoc"] = shutil.which("pandoc") is not None
        
        # Check Docling
        try:
            import docling
            availability["docling"] = True
        except ImportError:
            availability["docling"] = False
        
        return availability
    
    def run_conversion(self, converter: str) -> ConversionResult:
        """Run a single conversion and collect metrics."""
        output_name = f"quality_{converter}"
        output_file = self.output_dir / f"{output_name}.md"
        
        # Build command
        cmd = [
            "office2md",
            str(self.input_file),
            "-o", str(output_file),
        ]
        cmd.extend(self.CONVERTERS.get(converter, []))
        
        if self.verbose:
            cmd.append("-v")
            print(f"\n{'='*60}")
            print(f"Running: {' '.join(cmd)}")
            print(f"{'='*60}")
        
        # Run conversion
        start_time = time.time()
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120
            )
            elapsed = time.time() - start_time
            
            if result.returncode != 0:
                return ConversionResult(
                    converter=converter,
                    success=False,
                    time_seconds=elapsed,
                    error=result.stderr or "Unknown error"
                )
            
            # Collect metrics
            conv_result = ConversionResult(
                converter=converter,
                success=True,
                time_seconds=elapsed,
                output_file=output_file,
                images_dir=self.output_dir / f"{output_name}_images"
            )
            
            self._collect_metrics(conv_result)
            self._analyze_image_placement(conv_result)
            self._detect_quality_issues(conv_result)
            return conv_result
            
        except subprocess.TimeoutExpired:
            return ConversionResult(
                converter=converter,
                success=False,
                time_seconds=120,
                error="Timeout (120s)"
            )
        except Exception as e:
            return ConversionResult(
                converter=converter,
                success=False,
                time_seconds=time.time() - start_time,
                error=str(e)
            )
    
    def _collect_metrics(self, result: ConversionResult):
        """Collect quality metrics from conversion output."""
        if not result.output_file or not result.output_file.exists():
            return
        
        content = result.output_file.read_text(encoding="utf-8")
        
        # File size
        result.file_size = result.output_file.stat().st_size
        
        # Image count
        if result.images_dir and result.images_dir.exists():
            result.image_count = len(list(result.images_dir.glob("*")))
        
        # Heading count and samples
        headings = re.findall(r"^(#{1,6}\s+.+)$", content, re.MULTILINE)
        result.heading_count = len(headings)
        result.headings = headings[:10]
        
        # Table count - improved detection
        # Count pipe table separators (| --- | --- |)
        pipe_tables = len(re.findall(r"^\|[\s\-:|]+\|$", content, re.MULTILINE))
        # Count grid table headers (+===+===+)
        grid_tables = len(re.findall(r"^\+[=+]+\+$", content, re.MULTILINE))
        result.table_count = max(pipe_tables, grid_tables)
        
        # If no separators found, try counting table blocks
        if result.table_count == 0:
            # Look for consecutive lines starting with |
            table_blocks = re.findall(r"(^\|.+\|$\n)+", content, re.MULTILINE)
            # Each block with 2+ lines is likely a table
            result.table_count = len([b for b in table_blocks if b.count('\n') >= 2])
        
        # Get table sample
        table_lines = re.findall(r"^\|.+\|$", content, re.MULTILINE)
        if table_lines:
            result.table_sample = "\n".join(table_lines[:5])
        
        # List count
        result.list_count = len(re.findall(r"^[\s]*[-*+]\s+", content, re.MULTILINE))
        result.list_count += len(re.findall(r"^[\s]*\d+\.\s+", content, re.MULTILINE))
        
        # Bold count
        result.bold_count = len(re.findall(r"\*\*[^*]+\*\*", content))
        
        # Link count
        result.link_count = len(re.findall(r"\[([^\]]+)\]\(([^)]+)\)", content))
    
    def _analyze_image_placement(self, result: ConversionResult):
        """Analyze image placement in the document."""
        if not result.output_file or not result.output_file.exists():
            return
        
        content = result.output_file.read_text(encoding="utf-8")
        lines = content.split('\n')
        
        in_table = False
        images_in_current_cell = 0
        current_table_line = 0
        
        for i, line in enumerate(lines):
            # Detect table boundaries
            if line.strip().startswith('|') and '|' in line[1:]:
                if not in_table:
                    in_table = True
                    current_table_line = i
                
                # Count images in this table row
                images_in_line = len(re.findall(r'!\[.*?\]\(.*?\)', line))
                if images_in_line > 0:
                    result.images_in_tables += images_in_line
                    
                    # Check for multiple images in same cell (potential misplacement)
                    cells = line.split('|')
                    for cell in cells:
                        cell_images = len(re.findall(r'!\[.*?\]\(.*?\)', cell))
                        if cell_images > 2:  # More than 2 images in one cell is suspicious
                            result.image_issues.append(ImageIssue(
                                issue_type='misplaced',
                                location=f"Table at line {current_table_line}, cell with {cell_images} images",
                                details=f"Cell content: {cell[:100]}..."
                            ))
            else:
                if in_table and line.strip() and not line.strip().startswith('|'):
                    in_table = False
                
                # Count images outside tables
                images_in_line = len(re.findall(r'!\[.*?\]\(.*?\)', line))
                if images_in_line > 0:
                    result.images_outside_tables += images_in_line
        
        # Detect image placeholders (Docling issue)
        result.image_placeholders = len(re.findall(r'<!--\s*image\s*-->', content))
        if result.image_placeholders > 0:
            result.image_issues.append(ImageIssue(
                issue_type='missing',
                location='Throughout document',
                details=f'{result.image_placeholders} image placeholders (<!-- image -->) found instead of actual images'
            ))
        
        # Detect metadata leaks (Pandoc issue)
        metadata_leaks = re.findall(r'\{width="[^"]+"\s*height="[^"]+"\}', content)
        result.image_metadata_leaks = len(metadata_leaks)
        if result.image_metadata_leaks > 0:
            result.image_issues.append(ImageIssue(
                issue_type='metadata_leak',
                location='After images',
                details=f'{result.image_metadata_leaks} image metadata leaks ({{width="..." height="..."}}) found'
            ))
        
        # Detect consecutive images (potential misplacement)
        consecutive_pattern = r'(!\[.*?\]\([^)]+\)\s*){3,}'
        consecutive_matches = re.finditer(consecutive_pattern, content)
        for match in consecutive_matches:
            images_count = len(re.findall(r'!\[.*?\]\(.*?\)', match.group()))
            if images_count >= 3:
                # Find line number
                line_num = content[:match.start()].count('\n') + 1
                result.consecutive_images.append((images_count, f"line {line_num}"))
                result.image_issues.append(ImageIssue(
                    issue_type='misplaced',
                    location=f"Line {line_num}",
                    details=f'{images_count} consecutive images - likely misplaced'
                ))
    
    def _detect_quality_issues(self, result: ConversionResult):
        """Detect various quality issues in the output."""
        if not result.output_file or not result.output_file.exists():
            return
        
        content = result.output_file.read_text(encoding="utf-8")
        
        # Table issues
        # Check for malformed tables (inconsistent column counts)
        table_blocks = re.findall(r'(\|[^\n]+\|\n)+', content)
        for table in table_blocks:
            lines = table.strip().split('\n')
            if len(lines) >= 2:
                col_counts = [line.count('|') for line in lines]
                if len(set(col_counts)) > 1:
                    result.table_issues.append(
                        f"Inconsistent column count: {col_counts[:5]}..."
                    )
        
        # Check for Pandoc grid table syntax (not standard markdown)
        if re.search(r'\+[-=]+\+', content):
            result.table_issues.append("Grid table syntax detected (not standard Markdown)")
        
        # Formatting issues
        # Broken bold (e.g., **text****text**)
        broken_bold = re.findall(r'\*{4,}', content)
        if broken_bold:
            result.formatting_issues.append(f"Broken bold markers: {len(broken_bold)} occurrences")
        
        # Unclosed formatting
        lines = content.split('\n')
        for i, line in enumerate(lines):
            # Count asterisks (should be even for proper formatting)
            asterisk_count = line.count('**')
            if asterisk_count % 2 != 0:
                result.formatting_issues.append(f"Unclosed bold at line {i+1}")
        
        # Empty table cells where images should be
        empty_cells_after_image_pattern = r'\|\s*\|\s*!\['
        if re.search(empty_cells_after_image_pattern, content):
            result.image_issues.append(ImageIssue(
                issue_type='misplaced',
                location='Table cells',
                details='Empty cells adjacent to image cells detected'
            ))
    
    def run_all(self) -> QualityReport:
        """Run all conversions and generate report."""
        report = QualityReport(input_file=self.input_file)
        availability = self.check_converter_availability()
        
        print(f"\n{'='*60}")
        print(f"Quality Check: {self.input_file.name}")
        print(f"{'='*60}")
        print(f"Output directory: {self.output_dir}")
        print()
        
        # Check availability
        print("Converter Availability:")
        for conv in self.converters:
            status = "✅ Available" if availability.get(conv, False) else "❌ Not available"
            print(f"  {conv}: {status}")
        print()
        
        # Run conversions
        for converter in self.converters:
            if not availability.get(converter, False):
                print(f"⏭️  Skipping {converter} (not available)")
                continue
            
            print(f"🔄 Running {converter}...", end=" ", flush=True)
            result = self.run_conversion(converter)
            
            if result.success:
                print(f"✅ Done ({result.time_seconds:.2f}s)")
            else:
                print(f"❌ Failed: {result.error}")
            
            report.add_result(result)
        
        return report
    
    def print_report(self, report: QualityReport):
        """Print formatted quality report."""
        print(f"\n{'='*60}")
        print("QUALITY REPORT")
        print(f"{'='*60}")
        print(f"Input: {report.input_file}")
        print(f"Input size: {report.input_file.stat().st_size:,} bytes")
        print()
        
        # Summary table
        print("=" * 100)
        print(f"{'Converter':<12} {'Status':<8} {'Time':<8} {'Size':<10} {'Images':<8} {'Headings':<10} {'Tables':<8} {'Issues':<10}")
        print("=" * 100)
        
        for name, result in report.results.items():
            if result.success:
                status = "✅"
                time_str = f"{result.time_seconds:.2f}s"
                size_str = f"{result.file_size:,}"
                images_str = str(result.image_count)
                headings_str = str(result.heading_count)
                tables_str = str(result.table_count)
                issues_count = len(result.image_issues) + len(result.table_issues) + len(result.formatting_issues)
                issues_str = str(issues_count) if issues_count > 0 else "0"
            else:
                status = "❌"
                time_str = "-"
                size_str = "-"
                images_str = "-"
                headings_str = "-"
                tables_str = "-"
                issues_str = "-"
            
            print(f"{name:<12} {status:<8} {time_str:<8} {size_str:<10} {images_str:<8} {headings_str:<10} {tables_str:<8} {issues_str:<10}")
        
        print("=" * 100)
        print()
        
        # Detailed sections
        self._print_image_analysis(report)
        self._print_quality_issues(report)
        self._print_headings_comparison(report)
        self._print_tables_comparison(report)
        self._print_recommendations(report)
    
    def _print_image_analysis(self, report: QualityReport):
        """Print image placement analysis."""
        print("\n🖼️  IMAGE PLACEMENT ANALYSIS")
        print("-" * 60)
        
        print(f"\n{'Converter':<12} {'Total':<8} {'In Tables':<12} {'Outside':<10} {'Placeholders':<14} {'Metadata':<10}")
        print("-" * 66)
        
        for name, result in report.results.items():
            if not result.success:
                continue
            
            print(f"{name:<12} {result.image_count:<8} {result.images_in_tables:<12} {result.images_outside_tables:<10} {result.image_placeholders:<14} {result.image_metadata_leaks:<10}")
        
        # Print specific issues
        for name, result in report.results.items():
            if not result.success or not result.image_issues:
                continue
            
            print(f"\n⚠️  {name.upper()} Image Issues:")
            for issue in result.image_issues[:5]:  # Show first 5
                print(f"  [{issue.issue_type.upper()}] {issue.location}")
                print(f"    → {issue.details[:80]}...")
            
            if len(result.image_issues) > 5:
                print(f"  ... and {len(result.image_issues) - 5} more issues")
    
    def _print_quality_issues(self, report: QualityReport):
        """Print detected quality issues."""
        print("\n⚠️  QUALITY ISSUES DETECTED")
        print("-" * 60)
        
        has_issues = False
        
        for name, result in report.results.items():
            if not result.success:
                continue
            
            issues = []
            
            if result.table_issues:
                issues.extend([f"📊 Table: {i}" for i in result.table_issues])
            
            if result.formatting_issues:
                issues.extend([f"✏️  Format: {i}" for i in result.formatting_issues])
            
            if issues:
                has_issues = True
                print(f"\n{name.upper()}:")
                for issue in issues[:10]:
                    print(f"  {issue}")
                if len(issues) > 10:
                    print(f"  ... and {len(issues) - 10} more")
        
        if not has_issues:
            print("  ✅ No major quality issues detected")
    
    def _print_headings_comparison(self, report: QualityReport):
        """Print headings comparison."""
        print("\n📑 HEADINGS COMPARISON")
        print("-" * 40)
        
        for name, result in report.results.items():
            if not result.success:
                continue
            
            print(f"\n{name.upper()} ({result.heading_count} headings):")
            for h in result.headings[:5]:
                # Truncate long headings
                display = h[:60] + "..." if len(h) > 60 else h
                print(f"  {display}")
    
    def _print_tables_comparison(self, report: QualityReport):
        """Print tables comparison."""
        print("\n📊 TABLES COMPARISON")
        print("-" * 40)
        
        for name, result in report.results.items():
            if not result.success or not result.table_sample:
                continue
            
            print(f"\n{name.upper()} ({result.table_count} tables):")
            # Show first few lines of table
            lines = result.table_sample.split("\n")[:4]
            for line in lines:
                # Truncate long lines
                display = line[:70] + "..." if len(line) > 70 else line
                print(f"  {display}")
    
    def _print_recommendations(self, report: QualityReport):
        """Print recommendations based on results."""
        print("\n💡 RECOMMENDATIONS")
        print("-" * 60)
        
        successful = {k: v for k, v in report.results.items() if v.success}
        
        if not successful:
            print("❌ No successful conversions to compare")
            return
        
        # Analyze issues per converter
        print("\n📋 CONVERTER ASSESSMENT:")
        
        for name, result in successful.items():
            issues = []
            strengths = []
            
            # Check image handling
            if result.image_placeholders > 0:
                issues.append(f"❌ Images not embedded ({result.image_placeholders} placeholders)")
            elif result.image_metadata_leaks > 0:
                issues.append(f"⚠️  Image metadata leaks ({result.image_metadata_leaks})")
            else:
                if result.image_count > 0:
                    strengths.append(f"✅ Images extracted ({result.image_count})")
            
            # Check image placement
            misplaced = [i for i in result.image_issues if i.issue_type == 'misplaced']
            if misplaced:
                issues.append(f"⚠️  Potential image misplacement ({len(misplaced)} locations)")
            
            # Check tables
            if result.table_issues:
                issues.append(f"⚠️  Table issues ({len(result.table_issues)})")
            elif result.table_count > 0:
                strengths.append(f"✅ Tables converted ({result.table_count})")
            
            # Check formatting
            if result.formatting_issues:
                issues.append(f"⚠️  Formatting issues ({len(result.formatting_issues)})")
            else:
                if result.bold_count > 0:
                    strengths.append(f"✅ Formatting preserved")
            
            print(f"\n  {name.upper()}:")
            for s in strengths:
                print(f"    {s}")
            for i in issues:
                print(f"    {i}")
        
        # Overall recommendation
        print("\n🏆 BEST CONVERTER BY USE CASE:")
        
        # Score each converter
        scores = {}
        for name, result in successful.items():
            score = 100
            
            # Deduct for issues
            score -= len(result.image_issues) * 5
            score -= len(result.table_issues) * 10
            score -= len(result.formatting_issues) * 3
            score -= result.image_placeholders * 8  # Docling penalty
            score -= result.image_metadata_leaks * 2  # Pandoc penalty
            
            # Bonus for content
            if result.image_count > 0 and result.image_placeholders == 0:
                score += 10
            if result.table_count > 0 and not result.table_issues:
                score += 10
            
            scores[name] = max(0, score)
        
        best = max(scores.items(), key=lambda x: x[1])
        print(f"  🥇 Overall best: {best[0]} (score: {best[1]}/100)")
        
        # Specific recommendations
        print("\n  📝 Specific use cases:")
        
        # Best for images
        for name, result in successful.items():
            if result.image_count > 0 and result.image_placeholders == 0 and not [i for i in result.image_issues if i.issue_type == 'misplaced']:
                print(f"    • Best for images: {name}")
                break
        else:
            print("    • Best for images: ⚠️  All converters have image issues")
        
        # Best for tables
        for name, result in successful.items():
            if result.table_count > 0 and not result.table_issues:
                print(f"    • Best for tables: {name}")
                break
        else:
            print("    • Best for tables: ⚠️  All converters have table issues")
        
        # Best for speed
        fastest = min(successful.items(), key=lambda x: x[1].time_seconds)
        print(f"    • Fastest: {fastest[0]} ({fastest[1].time_seconds:.2f}s)")


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Quality check for office2md conversions",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python scripts/quality_check.py examples/input/doc1.docx
    python scripts/quality_check.py document.docx --output-dir ./tests
    python scripts/quality_check.py document.docx --converters default,pandoc
    python scripts/quality_check.py document.docx -v
        """
    )
    
    parser.add_argument(
        "input_file",
        help="Input file to convert and check"
    )
    
    parser.add_argument(
        "--output-dir", "-o",
        default="./quality_check_output",
        help="Output directory for converted files (default: ./quality_check_output)"
    )
    
    parser.add_argument(
        "--converters", "-c",
        default="default,pandoc,docling",
        help="Comma-separated list of converters to test (default: default,pandoc,docling)"
    )
    
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose output"
    )
    
    args = parser.parse_args()
    
    # Parse converters
    converters = [c.strip() for c in args.converters.split(",")]
    
    try:
        checker = QualityChecker(
            input_file=args.input_file,
            output_dir=args.output_dir,
            converters=converters,
            verbose=args.verbose
        )
        
        report = checker.run_all()
        checker.print_report(report)
        
        # Return success if at least one conversion succeeded
        success_count = sum(1 for r in report.results.values() if r.success)
        return 0 if success_count > 0 else 1
        
    except FileNotFoundError as e:
        print(f"❌ Error: {e}")
        return 1
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())