#!/usr/bin/env python3
"""
Demo: Using Extracted Excel Data with LLMs

This script demonstrates how to use the extracted JSON data with various LLM APIs
for analysis and insights generation.
"""

import json
from pathlib import Path
from typing import Dict, Any, List
import sys

class LLMDemo:
    """Demonstration class for LLM usage with extracted Excel data."""
    
    def __init__(self, json_file_path: Path):
        """Initialize with extracted JSON data."""
        self.json_file_path = json_file_path
        self.data = self._load_json_data()
    
    def _load_json_data(self) -> Dict[str, Any]:
        """Load the extracted JSON data."""
        with open(self.json_file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def generate_openai_prompt(self) -> str:
        """Generate a prompt suitable for OpenAI GPT models."""
        return f"""
You are an expert Excel analyst. Analyze this Excel workbook data and provide insights:

**Workbook Information:**
- File: {self.data['metadata']['filename']}
- Size: {self.data['metadata']['file_size_kb']} KB
- Sheets: {', '.join(self.data['metadata']['sheet_names'])}
- Total cells: {self.data['summary']['total_cells_with_data']:,}
- Total formulas: {self.data['summary']['total_formulas']:,}
- Complexity score: {self.data['summary']['complexity_score']}

**Sheet Analysis:**
{self._format_sheets_for_llm()}

**Key Data Types:**
{self._format_data_types_for_llm()}

**Formula Analysis:**
{self._format_formulas_for_llm()}

**Sample Data:**
{self._format_sample_data_for_llm()}

Please provide:
1. **Executive Summary**: What is this Excel file for?
2. **Structure Analysis**: How is it organized?
3. **Business Logic**: What calculations and processes does it represent?
4. **Key Insights**: What are the most important findings?
5. **Recommendations**: How could this be improved or used?

Format your response with clear sections and bullet points.
"""
    
    def generate_anthropic_prompt(self) -> str:
        """Generate a prompt suitable for Anthropic Claude models."""
        return f"""
<excel_workbook>
<metadata>
<filename>{self.data['metadata']['filename']}</filename>
<size_kb>{self.data['metadata']['file_size_kb']}</size_kb>
<sheets>{', '.join(self.data['metadata']['sheet_names'])}</sheets>
<total_cells>{self.data['summary']['total_cells_with_data']:,}</total_cells>
<total_formulas>{self.data['summary']['total_formulas']:,}</total_formulas>
<complexity_score>{self.data['summary']['complexity_score']}</complexity_score>
</metadata>

<structure>
{self._format_xml_structure()}
</structure>

<analysis>
{self._format_xml_analysis()}
</analysis>
</excel_workbook>

You are an expert Excel analyst. Analyze this Excel workbook and provide:

1. **Purpose**: What business function does this serve?
2. **Architecture**: How is it structured and organized?
3. **Logic**: What calculations and processes are implemented?
4. **Insights**: What are the key findings and patterns?
5. **Recommendations**: How could this be improved or optimized?

Provide a comprehensive analysis with specific examples from the data.
"""
    
    def generate_google_prompt(self) -> str:
        """Generate a prompt suitable for Google Gemini models."""
        return f"""
Analyze this Excel workbook data for business intelligence insights:

**Workbook Overview:**
- File: {self.data['metadata']['filename']}
- Complexity: {self._get_complexity_description()}
- Data Volume: {self.data['summary']['total_cells_with_data']:,} cells
- Calculations: {self.data['summary']['total_formulas']:,} formulas

**Sheet Breakdown:**
{self._format_sheets_for_llm()}

**Data Composition:**
{self._format_data_types_for_llm()}

**Key Calculations:**
{self._format_formulas_for_llm()}

**Business Context:**
Based on the data structure and calculations, this appears to be a {self._infer_business_type()}.

Please analyze and provide:
1. **Business Purpose**: What does this model/process represent?
2. **Data Flow**: How does information move through the sheets?
3. **Key Metrics**: What are the important calculations and outputs?
4. **Risk Assessment**: What potential issues or improvements exist?
5. **Actionable Insights**: What recommendations would you make?

Focus on practical business insights and actionable recommendations.
"""
    
    def _format_sheets_for_llm(self) -> str:
        """Format sheet information for LLM prompts."""
        lines = []
        for sheet_name, sheet_data in self.data['sheets'].items():
            summary = sheet_data['summary']
            lines.append(f"- **{sheet_name}**: {summary['total_cells_with_data']:,} cells, "
                        f"{summary['total_formulas']:,} formulas, "
                        f"{summary['total_tables']} tables")
        return '\n'.join(lines)
    
    def _format_data_types_for_llm(self) -> str:
        """Format data types for LLM prompts."""
        lines = []
        for data_type, count in self.data['summary']['data_types_summary'].items():
            percentage = (count / self.data['summary']['total_cells_with_data']) * 100
            lines.append(f"- {data_type}: {count:,} cells ({percentage:.1f}%)")
        return '\n'.join(lines)
    
    def _format_formulas_for_llm(self) -> str:
        """Format formula information for LLM prompts."""
        lines = []
        if self.data['summary']['formula_functions_summary']:
            for func, count in self.data['summary']['formula_functions_summary'].items():
                percentage = (count / self.data['summary']['total_formulas']) * 100
                lines.append(f"- {func}: {count} uses ({percentage:.1f}%)")
        else:
            lines.append("- No formulas found")
        return '\n'.join(lines)
    
    def _format_sample_data_for_llm(self) -> str:
        """Format sample data for LLM prompts."""
        lines = []
        for sheet_name, sheet_data in self.data['sheets'].items():
            if sheet_data['data']:
                lines.append(f"\n**{sheet_name}:**")
                count = 0
                for coord, cell_info in sheet_data['data'].items():
                    if count >= 3:
                        break
                    value = str(cell_info['value'])[:50]
                    lines.append(f"- {coord}: {value}")
                    count += 1
        return '\n'.join(lines)
    
    def _format_xml_structure(self) -> str:
        """Format structure information in XML for Claude."""
        lines = []
        for sheet_name, sheet_data in self.data['sheets'].items():
            summary = sheet_data['summary']
            lines.append(f"<sheet name='{sheet_name}'>")
            lines.append(f"  <cells>{summary['total_cells_with_data']}</cells>")
            lines.append(f"  <formulas>{summary['total_formulas']}</formulas>")
            lines.append(f"  <tables>{summary['total_tables']}</tables>")
            lines.append(f"</sheet>")
        return '\n'.join(lines)
    
    def _format_xml_analysis(self) -> str:
        """Format analysis information in XML for Claude."""
        lines = []
        lines.append(f"<data_types>")
        for data_type, count in self.data['summary']['data_types_summary'].items():
            lines.append(f"  <type name='{data_type}' count='{count}'/>")
        lines.append(f"</data_types>")
        
        lines.append(f"<formulas>")
        for func, count in self.data['summary']['formula_functions_summary'].items():
            lines.append(f"  <function name='{func}' count='{count}'/>")
        lines.append(f"</formulas>")
        
        return '\n'.join(lines)
    
    def _get_complexity_description(self) -> str:
        """Get a human-readable complexity description."""
        score = self.data['summary']['complexity_score']
        if score < 100:
            return "Simple"
        elif score < 500:
            return "Moderate"
        else:
            return "Complex"
    
    def _infer_business_type(self) -> str:
        """Infer the business type based on data patterns."""
        if self.data['summary']['total_tables'] > 0:
            return "data management system"
        elif 'VLOOKUP' in self.data['summary']['formula_functions_summary']:
            return "lookup and reference system"
        elif self.data['summary']['total_formulas'] > 50:
            return "financial or analytical model"
        else:
            return "basic data file"
    
    def generate_api_examples(self) -> Dict[str, str]:
        """Generate example API calls for different LLM providers."""
        examples = {}
        
        # OpenAI GPT-4 Example
        examples['openai'] = f"""
# OpenAI GPT-4 Example
import openai

client = openai.OpenAI(api_key="your-api-key")

response = client.chat.completions.create(
    model="gpt-4",
    messages=[
        {{"role": "system", "content": "You are an expert Excel analyst."}},
        {{"role": "user", "content": "{self.generate_openai_prompt()}"}}
    ],
    max_tokens=2000,
    temperature=0.3
)

print(response.choices[0].message.content)
"""
        
        # Anthropic Claude Example
        examples['anthropic'] = f"""
# Anthropic Claude Example
import anthropic

client = anthropic.Anthropic(api_key="your-api-key")

response = client.messages.create(
    model="claude-3-sonnet-20240229",
    max_tokens=2000,
    messages=[
        {{"role": "user", "content": "{self.generate_anthropic_prompt()}"}}
    ]
)

print(response.content[0].text)
"""
        
        # Google Gemini Example
        examples['gemini'] = f"""
# Google Gemini Example
import google.generativeai as genai

genai.configure(api_key="your-api-key")
model = genai.GenerativeModel('gemini-pro')

response = model.generate_content(
    "{self.generate_google_prompt()}"
)

print(response.text)
"""
        
        return examples
    
    def save_demo_files(self, output_dir: Path = None):
        """Save all demo files for easy testing."""
        if output_dir is None:
            output_dir = self.json_file_path.parent
        
        # Save prompts
        prompts_file = output_dir / f"{self.json_file_path.stem}_demo_prompts.md"
        with open(prompts_file, 'w', encoding='utf-8') as f:
            f.write(f"# LLM Demo Prompts for {self.data['metadata']['filename']}\n\n")
            
            f.write("## OpenAI GPT-4 Prompt\n\n")
            f.write("```\n")
            f.write(self.generate_openai_prompt())
            f.write("\n```\n\n")
            
            f.write("## Anthropic Claude Prompt\n\n")
            f.write("```\n")
            f.write(self.generate_anthropic_prompt())
            f.write("\n```\n\n")
            
            f.write("## Google Gemini Prompt\n\n")
            f.write("```\n")
            f.write(self.generate_google_prompt())
            f.write("\n```\n\n")
        
        # Save API examples
        api_file = output_dir / f"{self.json_file_path.stem}_api_examples.py"
        examples = self.generate_api_examples()
        
        with open(api_file, 'w', encoding='utf-8') as f:
            f.write('"""\n')
            f.write(f'API Examples for {self.data["metadata"]["filename"]}\n')
            f.write('"""\n\n')
            
            for provider, example in examples.items():
                f.write(example)
                f.write('\n\n')
        
        print(f"Demo files saved:")
        print(f"  - Prompts: {prompts_file}")
        print(f"  - API Examples: {api_file}")
        
        return prompts_file, api_file
    
    def print_quick_analysis(self):
        """Print a quick analysis to show what an LLM could provide."""
        print("=" * 80)
        print("QUICK LLM ANALYSIS DEMO")
        print("=" * 80)
        print()
        
        print("🤖 **What an LLM would analyze:**")
        print()
        
        # Business purpose inference
        print("📊 **Business Purpose**")
        if self.data['summary']['total_tables'] > 0:
            print("This appears to be a structured data management system with formal tables.")
        elif 'VLOOKUP' in self.data['summary']['formula_functions_summary']:
            print("This is a lookup and reference system, likely for product or inventory management.")
        elif self.data['summary']['total_formulas'] > 20:
            print("This is a calculation-heavy model, possibly for financial analysis or forecasting.")
        else:
            print("This is a basic data file for simple record keeping.")
        
        print()
        print("🏗️ **Architecture Analysis**")
        print(f"- Uses {self.data['metadata']['sheet_count']} sheets for separation of concerns")
        print(f"- Contains {self.data['summary']['total_formulas']:,} active calculations")
        
        if self.data['summary']['total_cross_sheet_references'] > 0:
            print(f"- Has {self.data['summary']['total_cross_sheet_references']} cross-sheet dependencies")
        
        print()
        print("🔍 **Key Insights**")
        
        # Data type insights
        if 'str' in self.data['summary']['data_types_summary']:
            str_percentage = (self.data['summary']['data_types_summary']['str'] / 
                            self.data['summary']['total_cells_with_data']) * 100
            if str_percentage > 50:
                print("- Heavy on text data, suggesting descriptive or categorical information")
        
        # Formula insights
        if self.data['summary']['formula_functions_summary']:
            most_common = max(self.data['summary']['formula_functions_summary'].items(), key=lambda x: x[1])
            print(f"- Primary calculation method: {most_common[0]} ({most_common[1]} instances)")
        
        print()
        print("💡 **LLM Recommendations**")
        
        complexity = self.data['summary']['complexity_score']
        if complexity < 100:
            print("- Simple file: Consider adding data validation for consistency")
        elif complexity < 500:
            print("- Moderate complexity: Document key formulas and assumptions")
        else:
            print("- High complexity: Implement comprehensive testing and documentation")
        
        print()
        print("=" * 80)


def main():
    """Main function to demonstrate LLM usage."""
    if len(sys.argv) < 2:
        print("Usage: python demo_llm_usage.py <json_file_path>")
        print("Example: python demo_llm_usage.py test_files/intermediate_model_data.json")
        return
    
    json_file_path = Path(sys.argv[1])
    if not json_file_path.exists():
        print(f"File not found: {json_file_path}")
        return
    
    # Create demo
    demo = LLMDemo(json_file_path)
    
    # Print quick analysis
    demo.print_quick_analysis()
    
    # Save demo files
    prompts_file, api_file = demo.save_demo_files()
    
    print(f"\n✅ Demo complete!")
    print(f"📄 Quick analysis printed above")
    print(f"📝 Prompts saved to: {prompts_file}")
    print(f"🔧 API examples saved to: {api_file}")
    print(f"\n🚀 You can now use these prompts with any LLM API!")


if __name__ == "__main__":
    main() 