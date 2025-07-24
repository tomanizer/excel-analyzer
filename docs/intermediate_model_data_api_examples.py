"""
API Examples for intermediate_model.xlsx
"""


# OpenAI GPT-4 Example
import openai

client = openai.OpenAI(api_key="your-api-key")

response = client.chat.completions.create(
    model="gpt-4",
    messages=[
        {"role": "system", "content": "You are an expert Excel analyst."},
        {"role": "user", "content": "
You are an expert Excel analyst. Analyze this Excel workbook data and provide insights:

**Workbook Information:**
- File: intermediate_model.xlsx
- Size: 6.92 KB
- Sheets: Data, Analysis, Summary
- Total cells: 66
- Total formulas: 17
- Complexity score: 135

**Sheet Analysis:**
- **Data**: 24 cells, 0 formulas, 1 tables
- **Analysis**: 37 cells, 15 formulas, 0 tables
- **Summary**: 5 cells, 2 formulas, 0 tables

**Key Data Types:**
- str: 41 cells (62.1%)
- int: 22 cells (33.3%)
- float: 3 cells (4.5%)

**Formula Analysis:**
- VLOOKUP: 10 uses (58.8%)
- SUM: 1 uses (5.9%)
- COUNT: 1 uses (5.9%)

**Sample Data:**

**Data:**
- A1: Product ID
- B1: Product Name
- C1: Price

**Analysis:**
- A1: Sales Analysis
- A3: Order ID
- B3: Product ID

**Summary:**
- A1: Sales Summary
- A3: Total Revenue
- B3: =SUM(Analysis!F4:F8)

Please provide:
1. **Executive Summary**: What is this Excel file for?
2. **Structure Analysis**: How is it organized?
3. **Business Logic**: What calculations and processes does it represent?
4. **Key Insights**: What are the most important findings?
5. **Recommendations**: How could this be improved or used?

Format your response with clear sections and bullet points.
"}
    ],
    max_tokens=2000,
    temperature=0.3
)

print(response.choices[0].message.content)



# Anthropic Claude Example
import anthropic

client = anthropic.Anthropic(api_key="your-api-key")

response = client.messages.create(
    model="claude-3-sonnet-20240229",
    max_tokens=2000,
    messages=[
        {"role": "user", "content": "
<excel_workbook>
<metadata>
<filename>intermediate_model.xlsx</filename>
<size_kb>6.92</size_kb>
<sheets>Data, Analysis, Summary</sheets>
<total_cells>66</total_cells>
<total_formulas>17</total_formulas>
<complexity_score>135</complexity_score>
</metadata>

<structure>
<sheet name='Data'>
  <cells>24</cells>
  <formulas>0</formulas>
  <tables>1</tables>
</sheet>
<sheet name='Analysis'>
  <cells>37</cells>
  <formulas>15</formulas>
  <tables>0</tables>
</sheet>
<sheet name='Summary'>
  <cells>5</cells>
  <formulas>2</formulas>
  <tables>0</tables>
</sheet>
</structure>

<analysis>
<data_types>
  <type name='str' count='41'/>
  <type name='int' count='22'/>
  <type name='float' count='3'/>
</data_types>
<formulas>
  <function name='VLOOKUP' count='10'/>
  <function name='SUM' count='1'/>
  <function name='COUNT' count='1'/>
</formulas>
</analysis>
</excel_workbook>

You are an expert Excel analyst. Analyze this Excel workbook and provide:

1. **Purpose**: What business function does this serve?
2. **Architecture**: How is it structured and organized?
3. **Logic**: What calculations and processes are implemented?
4. **Insights**: What are the key findings and patterns?
5. **Recommendations**: How could this be improved or optimized?

Provide a comprehensive analysis with specific examples from the data.
"}
    ]
)

print(response.content[0].text)



# Google Gemini Example
import google.generativeai as genai

genai.configure(api_key="your-api-key")
model = genai.GenerativeModel('gemini-pro')

response = model.generate_content(
    "
Analyze this Excel workbook data for business intelligence insights:

**Workbook Overview:**
- File: intermediate_model.xlsx
- Complexity: Moderate
- Data Volume: 66 cells
- Calculations: 17 formulas

**Sheet Breakdown:**
- **Data**: 24 cells, 0 formulas, 1 tables
- **Analysis**: 37 cells, 15 formulas, 0 tables
- **Summary**: 5 cells, 2 formulas, 0 tables

**Data Composition:**
- str: 41 cells (62.1%)
- int: 22 cells (33.3%)
- float: 3 cells (4.5%)

**Key Calculations:**
- VLOOKUP: 10 uses (58.8%)
- SUM: 1 uses (5.9%)
- COUNT: 1 uses (5.9%)

**Business Context:**
Based on the data structure and calculations, this appears to be a data management system.

Please analyze and provide:
1. **Business Purpose**: What does this model/process represent?
2. **Data Flow**: How does information move through the sheets?
3. **Key Metrics**: What are the important calculations and outputs?
4. **Risk Assessment**: What potential issues or improvements exist?
5. **Actionable Insights**: What recommendations would you make?

Focus on practical business insights and actionable recommendations.
"
)

print(response.text)


