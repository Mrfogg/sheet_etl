from autogen_agentchat.agents import AssistantAgent
from autogen_agentchat.conditions import TextMentionTermination
from autogen_agentchat.teams import RoundRobinGroupChat
from autogen_agentchat.ui import Console
from autogen_ext.models.openai import OpenAIChatCompletionClient
from autogen_core.tools import FunctionTool
import asyncio
import json
import csv
from tools import _generate_sheets_markdown_summary

class TimeSemantics:
    def __init__(self, data):
        self.has_time_dimension = data.get('has_time_dimension', False)
        self.time_columns = data.get('time_columns', [])
        self.time_granularity = data.get('time_granularity', '')
        self.likely_time_range = data.get('likely_time_range', '')

class RetrievalMetadata:
    def __init__(self, data):
        self.primary_topics = data.get('primary_topics', [])
        self.keywords = data.get('keywords', [])
        self.synonyms_or_aliases = data.get('synonyms_or_aliases', [])
        self.related_concepts = data.get('related_concepts', [])
        self.possible_user_queries = data.get('possible_user_queries', [])

class Confidence:
    def __init__(self, data):
        self.score = data.get('score', 0.0)
        self.notes = data.get('notes', '')
        self.ambiguities = data.get('ambiguities', [])

class ExcelSemanticUnderstanding:
    def __init__(self, json_data):
        self.file_id = json_data.get('file_id', '')
        self.file_name = json_data.get('file_name', '')
        self.sheet_name = json_data.get('sheet_name', '')
        self.semantic_title = json_data.get('semantic_title', '')
        self.semantic_summary = json_data.get('semantic_summary', '')
        self.business_domain = json_data.get('business_domain', [])
        self.data_purpose = json_data.get('data_purpose', [])
        self.data_category = json_data.get('data_category', [])
        self.row_granularity = json_data.get('row_granularity', '')
        self.core_entities = json_data.get('core_entities', [])
        self.metrics = json_data.get('metrics', [])
        self.dimensions = json_data.get('dimensions', [])
        self.time_semantics = TimeSemantics(json_data.get('time_semantics', {}))
        self.retrieval_metadata = RetrievalMetadata(json_data.get('retrieval_metadata', {}))
        self.confidence = Confidence(json_data.get('confidence', {}))

    def to_csv_row(self):
        return {
            'file_id': self.file_id,
            'file_name': self.file_name,
            'sheet_name': self.sheet_name,
            'semantic_title': self.semantic_title,
            'semantic_summary': self.semantic_summary,
            'business_domain': '|'.join(self.business_domain),
            'data_purpose': '|'.join(self.data_purpose),
            'data_category': '|'.join(self.data_category),
            'row_granularity': self.row_granularity,
            'core_entities': '|'.join(self.core_entities),
            'metrics': '|'.join(self.metrics),
            'dimensions': '|'.join(self.dimensions),
            'has_time_dimension': str(self.time_semantics.has_time_dimension),
            'time_columns': '|'.join(self.time_semantics.time_columns),
            'time_granularity': self.time_semantics.time_granularity,
            'likely_time_range': self.time_semantics.likely_time_range,
            'primary_topics': '|'.join(self.retrieval_metadata.primary_topics),
            'keywords': '|'.join(self.retrieval_metadata.keywords),
            'synonyms_or_aliases': '|'.join(self.retrieval_metadata.synonyms_or_aliases),
            'related_concepts': '|'.join(self.retrieval_metadata.related_concepts),
            'possible_user_queries': '|'.join(self.retrieval_metadata.possible_user_queries),
            'confidence_score': self.confidence.score,
            'confidence_notes': self.confidence.notes,
            'confidence_ambiguities': '|'.join(self.confidence.ambiguities)
        }

import os

def write_json_to_csv(json_data_list, output_file='semantic_understanding.csv'):
    if not json_data_list:
        print("No data to write.")
        return

    # Create objects from JSON data
    objects = [ExcelSemanticUnderstanding(json_data) for json_data in json_data_list]

    # Get header from the first object
    header = list(objects[0].to_csv_row().keys())

    # Check if file exists
    file_exists = os.path.exists(output_file)

    # Determine mode: write for new file, append for existing
    mode = 'a' if file_exists else 'w'

    with open(output_file, mode, newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=header)
        
        # Write header only if file is new
        if not file_exists:
            writer.writeheader()
        
        # Write data rows
        for obj in objects:
            writer.writerow(obj.to_csv_row())

    print(f"Data {'appended to' if file_exists else 'written to'} {output_file}")

model_client = OpenAIChatCompletionClient(
    model="doubao-seed-2-0-mini-260215",
    api_key="28f93d70-ea9c-47f8-8cce-5eae9e8cf3bc",
    base_url="https://ark.cn-beijing.volces.com/api/v3",
    model_info={
        "vision": False,
        "function_calling": True,
        "json_output": True,
        "family": "unknown",
        "structured_output": True
    },
)

system_prompt = """You are an Excel Semantic Understanding Agent.

Input:
You receive a structured representation of one Excel file, which may include:
- file_id
- file_name
- sheet_name
- row_count
- column_count
- headers
- markdown_table
- sample_rows
- optional metadata

Task:
Analyze the Excel representation and produce a structured semantic understanding that can be used by a downstream retrieval agent to match user questions with relevant Excel files.

Your semantic output must help answer:
- what this file is about
- what business domain it belongs to
- what one row represents
- what entities, metrics, and dimensions it contains
- what time semantics it has
- what user questions it can answer
- what keywords should be indexed for retrieval

Instructions:
- Use all available clues: file name, headers, table content, sample values, row/column counts, and metadata.
- Infer business meaning, not just literal table structure.
- Be retrieval-oriented.
- Be precise, compact, and evidence-based.
- Do not hallucinate unsupported details.
- If uncertain, explicitly say so.
- If the file contains aggregated data, identify it as summary/report data.
- If the file contains row-level records, identify the row granularity.
- If the file appears to be a reference table, mapping table, or master data table, identify that clearly.

Return JSON only using this schema:

{
  "file_id": "",
  "file_name": "",
  "sheet_name": "",
  "semantic_title": "",
  "semantic_summary": "",
  "business_domain": [],
  "data_purpose": [],
  "data_category": [],
  "row_granularity": "",
  "core_entities": [],
  "metrics": [],
  "dimensions": [],
  "time_semantics": {
    "has_time_dimension": false,
    "time_columns": [],
    "time_granularity": "",
    "likely_time_range": ""
  },
  "retrieval_metadata": {
    "primary_topics": [],
    "keywords": [],
    "synonyms_or_aliases": [],
    "related_concepts": [],
    "possible_user_queries": []
  },
  "confidence": {
    "score": 0.0,
    "notes": "",
    "ambiguities": []
  }
}"""

user_prompt = """
Analyze the following parsed Excel sheet and produce a semantic understanding according to the system prompt.

The input is a workbook preview exported into text. It may include:
- workbook-level metadata
- sheet-level metadata
- dimensions
- cell-address-style content (for example A1:, B2:, C3:)
- multi-row headers
- blank rows
- note rows
- trailing empty rows

Excel Input:
%s

Requirements:
- Treat cell addresses such as A1, B1, C1 as positional metadata, not business meaning.
- Detect the actual table structure.
- Reconstruct logical headers when header meaning is split across multiple rows.
- Separate the sheet into structural regions when possible:
  - title area
  - header area
  - data area
  - notes area
  - empty area
- Identify where real data rows begin and end.
- Ignore empty trailing rows for semantic interpretation.
- Do not mistake note rows for data rows.
- Infer the business meaning of the sheet based primarily on the table content.

Focus on:
- semantic title
- one-sentence summary
- business domain
- data purpose
- data category
- row granularity
- core entities
- metrics
- dimensions
- time semantics
- retrieval keywords
- synonyms / aliases
- possible user questions
- confidence and ambiguity

Return JSON only.
Follow the required schema exactly.
"""

excel_summary_agent = AssistantAgent(
    "summary_agent",
    model_client=model_client,
    description="Excel Semantic Understanding Agent",
    system_message=system_prompt,
    model_client_stream=True
)

import os
import glob

async def process_excel_file(excel_path):
    """处理单个Excel文件"""
    print(f"\n🔍 处理文件: {os.path.basename(excel_path)}")
    
    # 生成Excel markdown摘要
    markdowns = _generate_sheets_markdown_summary(excel_path,sheet_index=0)
    
    # 调用智能体处理
    for markdown in markdowns:
        result = await excel_summary_agent.run(task=user_prompt % markdown)
    
        # 获取最终结果
        final_content = result.messages[-1].content
        print("📊 最终总结结果：")
        print(final_content)
        
        # 尝试解析JSON并写入CSV
        try:
            json_data = json.loads(final_content)
            write_json_to_csv([json_data])
        except json.JSONDecodeError as e:
            print(f"Error parsing JSON: {e}")

async def main() -> None:
    # 关闭 AutoGen 自动打印（关键！）
    import logging
    logging.disable(logging.INFO)  # 屏蔽中间过程输出
    
    # 获取excel_base目录下的所有Excel文件
    excel_files = glob.glob('excel_base/*.xlsx')
    
    if not excel_files:
        print("没有找到Excel文件")
        return
    
    print(f"找到 {len(excel_files)} 个Excel文件")
    
    # 处理每个Excel文件
    for excel_file in excel_files:
        await process_excel_file(excel_file)

asyncio.run(main())
