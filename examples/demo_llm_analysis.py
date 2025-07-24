#!/usr/bin/env python3
"""
Test LLM Analysis with Extracted Excel Data

This script demonstrates how to use the extracted JSON data from Excel files
with an LLM for analysis, summarization, and insights generation.
"""

import json
from pathlib import Path
from typing import Dict, Any, List
import sys

class LLMAnalysisTester:
    """Test class for LLM analysis of extracted Excel data."""
    
    def __init__(self, json_file_path: Path):
        """Initialize with extracted JSON data."""
        self.json_file_path = json_file_path
        self.data = self._load_json_data()
    
    def _load_json_data(self) -> Dict[str, Any]:
        """Load the extracted JSON data."""
        with open(self.json_file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def generate_analysis_prompts(self) -> List[Dict[str, str]]:
        """Generate various analysis prompts for LLM testing."""
        
        prompts = []
        
        # 1. Basic Summary Prompt
        prompts.append({
            'name': 'Basic Summary',
            'prompt': f"""
Analyze this Excel workbook data and provide a comprehensive summary:

**File Information:**
- Filename: {self.data['metadata']['filename']}
- Size: {self.data['metadata']['file_size_kb']} KB
- Sheets: {self.data['metadata']['sheet_count']}
- Sheet Names: {', '.join(self.data['metadata']['sheet_names'])}

**Key Statistics:**
- Total cells with data: {self.data['summary']['total_cells_with_data']:,}
- Total formulas: {self.data['summary']['total_formulas']:,}
- Total tables: {self.data['summary']['total_tables']}
- Total charts: {self.data['summary']['total_charts']}
- Named ranges: {self.data['summary']['total_named_ranges']}
- Cross-sheet references: {self.data['summary']['total_cross_sheet_references']}

**Complexity Score:** {self.data['summary']['complexity_score']}

Please provide:
1. A brief overview of what this Excel file appears to be for
2. The main types of data and calculations it contains
3. Key insights about its structure and complexity
4. Recommendations for understanding or working with this file
"""
        })
        
        # 2. Structure Analysis Prompt
        prompts.append({
            'name': 'Structure Analysis',
            'prompt': f"""
Analyze the structure of this Excel workbook and explain how it's organized:

**Sheet Structure:**
{self._format_sheet_structure()}

**Data Types Found:**
{self._format_data_types()}

**Formula Analysis:**
{self._format_formula_analysis()}

Please provide:
1. How the workbook is structured across different sheets
2. What each sheet appears to be responsible for
3. How data flows between sheets (if any)
4. The overall architecture and design patterns used
"""
        })
        
        # 3. Business Logic Analysis Prompt
        prompts.append({
            'name': 'Business Logic Analysis',
            'prompt': f"""
Analyze the business logic and calculations in this Excel workbook:

**Key Calculations:**
{self._format_key_calculations()}

**Data Relationships:**
{self._format_relationships()}

**Tables and Structured Data:**
{self._format_tables_info()}

Please provide:
1. What business processes or models this Excel file represents
2. The main calculations and their business purpose
3. Key assumptions and variables used
4. Potential business insights that could be derived
"""
        })
        
        # 4. Data Quality Assessment Prompt
        prompts.append({
            'name': 'Data Quality Assessment',
            'prompt': f"""
Assess the data quality and structure of this Excel workbook:

**Data Distribution:**
{self._format_data_distribution()}

**Formula Complexity:**
{self._format_formula_complexity()}

**Validation and Controls:**
{self._format_validation_info()}

Please provide:
1. Assessment of data quality and consistency
2. Potential data integrity issues
3. Recommendations for data validation
4. Suggestions for improving data structure
"""
        })
        
        # 5. Migration/Conversion Analysis Prompt
        prompts.append({
            'name': 'Migration Analysis',
            'prompt': f"""
Analyze this Excel workbook for potential migration to other systems:

**Current Structure:**
- Complexity Score: {self.data['summary']['complexity_score']}
- Total Formulas: {self.data['summary']['total_formulas']:,}
- Cross-sheet References: {self.data['summary']['total_cross_sheet_references']}
- External Dependencies: {len(self.data['global_features']['external_links'])}

**Technical Features:**
{self._format_technical_features()}

Please provide:
1. Assessment of migration complexity
2. Recommended migration approach
3. Potential challenges and risks
4. Suggested target systems or platforms
5. Estimated effort and timeline
"""
        })
        
        return prompts
    
    def _format_sheet_structure(self) -> str:
        """Format sheet structure information."""
        lines = []
        for sheet_name, sheet_data in self.data['sheets'].items():
            summary = sheet_data['summary']
            lines.append(f"- **{sheet_name}**: {summary['total_cells_with_data']:,} cells, "
                        f"{summary['total_formulas']:,} formulas, "
                        f"{summary['total_tables']} tables, "
                        f"{summary['total_charts']} charts")
        return '\n'.join(lines)
    
    def _format_data_types(self) -> str:
        """Format data types information."""
        lines = []
        for data_type, count in self.data['summary']['data_types_summary'].items():
            percentage = (count / self.data['summary']['total_cells_with_data']) * 100
            lines.append(f"- {data_type}: {count:,} cells ({percentage:.1f}%)")
        return '\n'.join(lines)
    
    def _format_formula_analysis(self) -> str:
        """Format formula analysis information."""
        lines = []
        for func, count in self.data['summary']['formula_functions_summary'].items():
            percentage = (count / self.data['summary']['total_formulas']) * 100
            lines.append(f"- {func}: {count} uses ({percentage:.1f}%)")
        return '\n'.join(lines)
    
    def _format_key_calculations(self) -> str:
        """Format key calculations information."""
        lines = []
        # Show sample formulas from each sheet
        for sheet_name, sheet_data in self.data['sheets'].items():
            if sheet_data['formulas']:
                lines.append(f"\n**{sheet_name}:**")
                count = 0
                for coord, formula_info in sheet_data['formulas'].items():
                    if count >= 3:  # Limit to 3 formulas per sheet
                        break
                    formula = formula_info['formula'][:100]  # Truncate long formulas
                    lines.append(f"- {coord}: {formula}")
                    count += 1
        return '\n'.join(lines)
    
    def _format_relationships(self) -> str:
        """Format relationships information."""
        lines = []
        cross_refs = self.data['relationships']['cross_sheet_references']
        if cross_refs:
            lines.append("**Cross-sheet References:**")
            for ref in cross_refs[:5]:  # Limit to 5 references
                lines.append(f"- {ref['source_sheet']}!{ref['source_cell']} → "
                           f"{ref['target_sheet']}!{ref['target_cell']}")
        else:
            lines.append("No cross-sheet references found")
        return '\n'.join(lines)
    
    def _format_tables_info(self) -> str:
        """Format tables information."""
        lines = []
        for sheet_name, sheet_data in self.data['sheets'].items():
            if sheet_data['tables']:
                lines.append(f"\n**{sheet_name}:**")
                for table in sheet_data['tables']:
                    lines.append(f"- {table['name']}: {table['range']}")
        return '\n'.join(lines) if lines else "No formal tables found"
    
    def _format_data_distribution(self) -> str:
        """Format data distribution information."""
        lines = []
        for sheet_name, sheet_data in self.data['sheets'].items():
            summary = sheet_data['summary']
            lines.append(f"- **{sheet_name}**: {summary['total_cells_with_data']:,} cells "
                        f"({summary['total_cells_with_data'] / self.data['summary']['total_cells_with_data'] * 100:.1f}%)")
        return '\n'.join(lines)
    
    def _format_formula_complexity(self) -> str:
        """Format formula complexity information."""
        lines = []
        total_formulas = self.data['summary']['total_formulas']
        if total_formulas > 0:
            lines.append(f"Total formulas: {total_formulas:,}")
            lines.append("Most complex functions:")
            sorted_funcs = sorted(self.data['summary']['formula_functions_summary'].items(), 
                                key=lambda x: x[1], reverse=True)[:5]
            for func, count in sorted_funcs:
                lines.append(f"- {func}: {count} uses")
        else:
            lines.append("No formulas found")
        return '\n'.join(lines)
    
    def _format_validation_info(self) -> str:
        """Format validation information."""
        lines = []
        total_validations = sum(sheet_data['summary']['total_validations'] 
                              for sheet_data in self.data['sheets'].values())
        if total_validations > 0:
            lines.append(f"Total validation rules: {total_validations}")
            for sheet_name, sheet_data in self.data['sheets'].items():
                if sheet_data['data_validations']:
                    lines.append(f"- {sheet_name}: {len(sheet_data['data_validations'])} rules")
        else:
            lines.append("No data validation rules found")
        return '\n'.join(lines)
    
    def _format_technical_features(self) -> str:
        """Format technical features information."""
        lines = []
        
        # Named ranges
        if self.data['global_features']['named_ranges']:
            lines.append("**Named Ranges:**")
            for name, destinations in self.data['global_features']['named_ranges'].items():
                lines.append(f"- {name}: {', '.join(destinations)}")
        
        # External links
        if self.data['global_features']['external_links']:
            lines.append("\n**External Links:**")
            for link in self.data['global_features']['external_links']:
                lines.append(f"- {link}")
        
        # Charts
        total_charts = sum(sheet_data['summary']['total_charts'] 
                          for sheet_data in self.data['sheets'].values())
        if total_charts > 0:
            lines.append(f"\n**Charts:** {total_charts} total")
        
        return '\n'.join(lines) if lines else "No advanced technical features found"
    
    def save_prompts_to_file(self, output_file: Path = None):
        """Save all prompts to a markdown file for easy testing."""
        if output_file is None:
            output_file = self.json_file_path.with_name(f"{self.json_file_path.stem}_llm_prompts.md")
        
        prompts = self.generate_analysis_prompts()
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(f"# LLM Analysis Prompts for {self.data['metadata']['filename']}\n\n")
            f.write("This file contains various prompts you can use to analyze the extracted Excel data with an LLM.\n\n")
            
            for i, prompt_data in enumerate(prompts, 1):
                f.write(f"## {i}. {prompt_data['name']}\n\n")
                f.write("**Prompt:**\n\n")
                f.write("```\n")
                f.write(prompt_data['prompt'].strip())
                f.write("\n```\n\n")
                f.write("---\n\n")
        
        print(f"Prompts saved to: {output_file}")
        return output_file
    
    def print_sample_analysis(self):
        """Print a sample analysis to demonstrate what an LLM could produce."""
        print("=" * 80)
        print("SAMPLE LLM ANALYSIS")
        print("=" * 80)
        print()
        
        # Basic analysis
        print("📊 **WORKBOOK OVERVIEW**")
        print(f"This Excel file '{self.data['metadata']['filename']}' appears to be a ")
        
        if self.data['summary']['total_tables'] > 0:
            print("structured data model with formal tables for data organization.")
        elif self.data['summary']['total_formulas'] > 100:
            print("complex financial or analytical model with extensive calculations.")
        elif self.data['summary']['total_charts'] > 0:
            print("reporting or dashboard file with visualizations.")
        else:
            print("simple data file or basic model.")
        
        print()
        print("📋 **STRUCTURE ANALYSIS**")
        print(f"- Contains {self.data['metadata']['sheet_count']} sheets: {', '.join(self.data['metadata']['sheet_names'])}")
        print(f"- Total data cells: {self.data['summary']['total_cells_with_data']:,}")
        print(f"- Active calculations: {self.data['summary']['total_formulas']:,} formulas")
        
        if self.data['summary']['total_cross_sheet_references'] > 0:
            print(f"- Interconnected: {self.data['summary']['total_cross_sheet_references']} cross-sheet references")
        
        print()
        print("🔍 **KEY INSIGHTS**")
        
        # Data type insights
        if 'str' in self.data['summary']['data_types_summary']:
            text_cells = self.data['summary']['data_types_summary']['str']
            if text_cells > self.data['summary']['total_cells_with_data'] * 0.3:
                print("- Contains significant text data (labels, descriptions, categories)")
        
        if 'float' in self.data['summary']['data_types_summary'] or 'int' in self.data['summary']['data_types_summary']:
            print("- Includes numerical data for calculations and analysis")
        
        # Formula insights
        if self.data['summary']['formula_functions_summary']:
            most_common = max(self.data['summary']['formula_functions_summary'].items(), key=lambda x: x[1])
            print(f"- Most common calculation: {most_common[0]} function ({most_common[1]} uses)")
        
        print()
        print("💡 **RECOMMENDATIONS**")
        
        complexity_score = self.data['summary']['complexity_score']
        if complexity_score < 100:
            print("- This is a simple file suitable for basic analysis or data entry")
        elif complexity_score < 500:
            print("- Moderate complexity - consider documenting key formulas and assumptions")
        else:
            print("- High complexity - recommend thorough documentation and testing")
        
        if self.data['summary']['total_cross_sheet_references'] > 10:
            print("- Multiple sheet dependencies - ensure data consistency across sheets")
        
        print()
        print("=" * 80)


def main():
    """Main function to test LLM analysis capabilities."""
    if len(sys.argv) < 2:
        print("Usage: python test_llm_analysis.py <json_file_path>")
        print("Example: python demo_llm_analysis.py reports/intermediate_model_data.json")
        return
    
    json_file_path = Path(sys.argv[1])
    if not json_file_path.exists():
        print(f"File not found: {json_file_path}")
        return
    
    # Create tester
    tester = LLMAnalysisTester(json_file_path)
    
    # Print sample analysis
    tester.print_sample_analysis()
    
    # Save prompts to file
    prompts_file = tester.save_prompts_to_file()
    
    print(f"\n✅ Analysis complete!")
    print(f"📄 Sample analysis printed above")
    print(f"📝 LLM prompts saved to: {prompts_file}")
    print(f"\n💡 You can now use these prompts with any LLM to analyze the Excel data!")


if __name__ == "__main__":
    main() 