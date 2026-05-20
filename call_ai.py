#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Scheduling AI Call Program
Read prompt files and Excel data, call AI to generate schedule
"""

import os
import configparser
import json
import pandas as pd
from datetime import datetime
from openai import OpenAI


def load_config(config_file='config.ini'):
    """Load configuration file"""
    config = configparser.ConfigParser()
    config.read(config_file, encoding='utf-8')

    return {
        'api_key': config.get('api', 'api_key'),
        'base_url': config.get('api', 'base_url'),
        'model': config.get('api', 'model'),
        'max_tokens': config.getint('api', 'max_tokens', fallback=8000),
        'temperature': config.getfloat('api', 'temperature', fallback=0.1),
        'timeout': config.getint('api', 'timeout', fallback=300)
    }


def load_prompt_parts():
    """Load three parts of prompts"""
    parts = {}

    # Part 1: Fixed rules
    part1_file = 'prompt_part1_fixed_rules.md'
    if not os.path.exists(part1_file):
        raise FileNotFoundError(f"Prompt file not found: {part1_file}")
    with open(part1_file, 'r', encoding='utf-8') as f:
        parts['part1'] = f.read()

    # Part 2: Variable parameters
    part2_file = 'prompt_part2_variables.md'
    if not os.path.exists(part2_file):
        raise FileNotFoundError(f"Prompt file not found: {part2_file}")
    with open(part2_file, 'r', encoding='utf-8') as f:
        parts['part2'] = f.read()

    # Part 3: Output rules
    part3_file = 'prompt_part3_output_rules.md'
    if not os.path.exists(part3_file):
        raise FileNotFoundError(f"Prompt file not found: {part3_file}")
    with open(part3_file, 'r', encoding='utf-8') as f:
        parts['part3'] = f.read()

    return parts


def extract_system_name(change_name, system_name):
    """Extract system name"""
    import re

    # Prioritize extracting system name from change name (find consecutive uppercase letters, 2 or more)
    if pd.notna(change_name) and str(change_name) not in ['nan', 'None', '']:
        match = re.match(r'^([A-Z]{2,})', str(change_name))
        if match:
            return match.group(1)

    # If system name exists and is not empty/nan, extract English letter part
    if pd.notna(system_name) and str(system_name) not in ['nan', 'None', '']:
        system_str = str(system_name).strip()
        # Extract English letter part at the beginning (e.g., "ELK-集中式日誌管理平台" -> "ELK")
        match = re.match(r'^([A-Z]{2,})', system_str)
        if match:
            return match.group(1)
        else:
            # If no match, try to extract part before hyphen
            if '-' in system_str:
                return system_str.split('-')[0].strip()
            return system_str

    return 'Unknown'


def format_submitter_name(submitter):
    """Format submitter name"""
    if pd.notna(submitter) and str(submitter) not in ['nan', 'None', '']:
        # Remove employee ID part (e.g., "譚翀/mo164291" -> "譚翀")
        submitter_str = str(submitter)
        if '/' in submitter_str:
            return submitter_str.split('/')[0].strip()
        return submitter_str.strip()
    return 'Unknown'


def load_excel_data(excel_file='data/sheet_input.xlsx'):
    """Read Excel file data"""
    if not os.path.exists(excel_file):
        raise FileNotFoundError(f"Excel file not found: {excel_file}")

    print(f"Reading Excel file: {excel_file}")
    df = pd.read_excel(excel_file)
    print(f"  ✓ Read {len(df)} rows of data")

    # Apply filters according to prompt_part2_variables.md
    print(f"\n[Filtering] Applying data filters...")

    # Filter 1: Remove cancelled work orders
    if '工单状态' in df.columns:
        before_count = len(df)
        df = df[df['工单状态'] != '已撤銷']
        after_count = len(df)
        filtered_count = before_count - after_count
        if filtered_count > 0:
            print(f"  ✓ Filtered {filtered_count} cancelled work orders")

    # Filter 2: Remove specific teams
    filter_teams = ['安全內控團隊', '前端網絡團隊', '系統平臺團隊']
    if '提單人所屬團隊' in df.columns:
        before_count = len(df)
        df = df[~df['提單人所屬團隊'].isin(filter_teams)]
        after_count = len(df)
        filtered_count = before_count - after_count
        if filtered_count > 0:
            print(f"  ✓ Filtered {filtered_count} work orders from specific teams")
            for team in filter_teams:
                team_count = len(df[df['提單人所屬團隊'] == team])
                if team_count > 0:
                    print(f"    - {team}: {team_count} work orders")

    print(f"  ✓ Final work order count: {len(df)}")

    # Convert to AI-friendly text format
    data_text = f"""
## Excel File Data Information
File: {excel_file}
Total rows: {len(df)}
Data source: Real work order data, must schedule strictly according to this data

## Work Order Details ({len(df)} work orders total)
"""

    # Generate work order information according to Part 1 format requirements
    for idx, row in df.iterrows():
        order_id = row.get('工單號', 'N/A')
        change_name = row.get('變更名稱', 'N/A')
        system_name_raw = row.get('變更系統名稱匯總', 'N/A')
        submitter_raw = row.get('提單人', 'N/A')
        start_time = row.get('計劃開始時間', 'N/A')
        end_time = row.get('計劃結束時間', 'N/A')

        # Extract the number after the last hyphen from work order ID
        if isinstance(order_id, str) and '-' in order_id:
            order_num = order_id.split('-')[-1]
        else:
            order_num = str(order_id)

        # Extract system name
        system_name = extract_system_name(change_name, system_name_raw)

        # Format submitter name
        submitter = format_submitter_name(submitter_raw)

        # Format time
        if pd.notna(start_time) and str(start_time) not in ['nan', 'None', '']:
            try:
                dt = pd.to_datetime(start_time)
                date_str = dt.strftime('%Y-%m-%d')
                time_str = dt.strftime('%H:%M')
            except:
                date_str = str(start_time)
                time_str = "Unknown time"
        else:
            date_str = "Unknown date"
            time_str = "Unknown time"

        # Display according to Part 1 format requirements
        data_text += f"\n### Work Order {idx + 1}: {order_num} {system_name} {submitter}\n"
        data_text += f"- Full work order ID: {order_id}\n"
        data_text += f"- Change name: {change_name}\n"
        data_text += f"- Planned start time: {start_time} (Date: {date_str}, Time: {time_str})\n"
        data_text += f"- Planned end time: {end_time}\n"

        # Add other important information
        if '工单状态' in df.columns:
            data_text += f"- Work order status: {row.get('工单状态', 'N/A')}\n"

        if '提單人所屬團隊' in df.columns:
            data_text += f"- Team: {row.get('提單人所屬團隊', 'N/A')}\n"

    # Add time analysis
    data_text += "\n## Time Distribution Analysis\n"

    # Extract all dates
    dates = []
    for idx, row in df.iterrows():
        start_time = row.get('計劃開始時間', 'N/A')
        if pd.notna(start_time) and str(start_time) not in ['nan', 'None', '']:
            try:
                dt = pd.to_datetime(start_time)
                date_key = dt.strftime('%Y-%m-%d')
                if date_key not in dates:
                    dates.append(date_key)
            except:
                pass

    if dates:
        dates.sort()
        data_text += f"- Dates included: {', '.join(dates)}\n"

        # Count work orders per day
        from collections import defaultdict
        date_counts = defaultdict(int)
        for idx, row in df.iterrows():
            start_time = row.get('計劃開始時間', 'N/A')
            if pd.notna(start_time) and str(start_time) not in ['nan', 'None', '']:
                try:
                    dt = pd.to_datetime(start_time)
                    date_key = dt.strftime('%Y-%m-%d')
                    date_counts[date_key] += 1
                except:
                    pass

        data_text += "- Daily work order distribution:\n"
        for date in sorted(date_counts.keys()):
            data_text += f"  * {date}: {date_counts[date]} work orders\n"

    data_text += f"\n## Important Reminders\n"
    data_text += f"- The above {len(df)} work orders are real scheduling data, all must be scheduled, no omissions\n"
    data_text += f"- System name extraction rule: If 'Change System Name Summary' is empty, extract English letters from the beginning of 'Change Name'\n"
    data_text += f"- Submitter name processing: Remove employee ID part, keep only name\n"
    data_text += f"- Work order ID format: Use the number after the last hyphen\n"

    return data_text, df


def combine_prompt(prompt_parts, excel_data_text):
    """Combine complete prompt"""
    full_prompt = f"""
# Complete Scheduling Task Prompt

{prompt_parts['part1']}

{prompt_parts['part2']}

## Excel File Data
{excel_data_text}

{prompt_parts['part3']}

## Important Reminders
1. You must schedule according to the real data in the above Excel file
2. Strictly follow the format requirements of Part 1 and personnel information of Part 2
3. Generate final results according to the output format of Part 3
4. Ensure all work orders are assigned, do not omit any data

Please start generating the schedule now, output in JSON format (for easy conversion to Excel).
"""
    return full_prompt


def call_ai(api_key, base_url, model, prompt, max_tokens=8000, temperature=0.7, timeout=300):
    """Call AI API"""
    print(f"Initializing AI client...")
    print(f"Model: {model}")
    print(f"API URL: {base_url}")
    print(f"Prompt length: {len(prompt)} characters")

    client = OpenAI(
        base_url=base_url,
        api_key=api_key,
        timeout=timeout
    )

    try:
        print(f"\nSending request to AI server (may take 1-3 minutes)...\n")

        completion = client.chat.completions.create(
            model=model,
            messages=[
                {
                    "role": "system",
                    "content": "You are a professional scheduling assistant. 🚨🚨🚨 HIGHEST PRIORITY RULES 🚨🚨🚨: 1) BALANCED ASSIGNMENT IS ABSOLUTE MANDATORY: Every person MUST receive at least 1 work order. Zero work order assignments = CRITICAL FAILURE. 2) Work order counts MUST be balanced (difference no more than 1 between people). 3) When multiple work orders exist at the same time, ONE PERSON can handle MULTIPLE work orders displayed with line breaks in the same cell. 4) Each work order appears exactly once globally. 5) Work orders must be assigned according to their planned start time date. If you cannot achieve balanced assignment with the given data, you MUST redistribute work orders until balance is achieved. NEVER output a schedule where some people have 0 work orders."
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            temperature=temperature,
            max_tokens=max_tokens,
            stream=False
        )

        print(f"Receiving AI response, processing...")

        # Get response content
        if hasattr(completion, 'choices') and len(completion.choices) > 0:
            message = completion.choices[0].message

            # Try different fields (some models use reasoning_content)
            content = None
            for field in ['content', 'reasoning_content', 'text', 'response']:
                if hasattr(message, field):
                    value = getattr(message, field)
                    if value:
                        content = value
                        print(f"  ✓ Found content in '{field}' field")
                        break

            if not content:
                raise ValueError("No content found in API response")

            print(f"API call successful!")
            print(f"Response length: {len(content)} characters")

            # Show token usage
            if hasattr(completion, 'usage') and completion.usage:
                usage = completion.usage
                print(f"Token usage: Input={usage.prompt_tokens}, Output={usage.completion_tokens}, Total={usage.total_tokens}")

            return content.strip()
        else:
            raise ValueError("API response format error")

    except Exception as e:
        print(f"Error calling AI: {type(e).__name__}: {str(e)}")
        raise


def save_response(content, output_file='ai_responses/schedule_result.json'):
    """Save AI response"""
    # Ensure directory exists
    os.makedirs(os.path.dirname(output_file) if os.path.dirname(output_file) else '.', exist_ok=True)

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"\nResponse saved to: {output_file}")

    return output_file


def main():
    """Main function"""
    print("=" * 60)
    print("Scheduling AI Call Program")
    print("=" * 60)
    print()

    try:
        # 1. Load configuration
        print("[1/5] Loading configuration file...")
        config = load_config()
        print(f"  ✓ Configuration loaded successfully")

        # 2. Load prompts
        print("\n[2/5] Loading prompt files...")
        prompt_parts = load_prompt_parts()
        print(f"  ✓ Part 1 (Fixed rules): {len(prompt_parts['part1'])} characters")
        print(f"  ✓ Part 2 (Variable parameters): {len(prompt_parts['part2'])} characters")
        print(f"  ✓ Part 3 (Output rules): {len(prompt_parts['part3'])} characters")

        # 3. Read Excel data
        print("\n[3/5] Reading Excel data...")
        excel_data_text, df = load_excel_data('data/sheet_input.xlsx')
        print(f"  ✓ Excel data processing completed")

        # 4. Combine prompts and call AI
        print("\n[4/5] Combining prompts and calling AI...")
        full_prompt = combine_prompt(prompt_parts, excel_data_text)
        print(f"  ✓ Complete prompt length: {len(full_prompt)} characters")

        response = call_ai(
            api_key=config['api_key'],
            base_url=config['base_url'],
            model=config['model'],
            prompt=full_prompt,
            max_tokens=config['max_tokens'],
            temperature=config['temperature'],
            timeout=config['timeout']
        )

        # 5. Save response
        print("\n[5/5] Saving AI response...")
        output_file = save_response(response)

        # Show response preview
        print("\n" + "=" * 60)
        print("AI Response Preview:")
        print("=" * 60)
        preview = response[:800] if len(response) > 800 else response
        print(preview)
        if len(response) > 800:
            print(f"... (Total {len(response)} characters)")

        print("\n" + "=" * 60)
        print("Schedule generation completed!")
        print(f"Results saved to: {output_file}")
        print("=" * 60)

        # Convert JSON to Excel
        print("\n[Extra] Converting JSON to Excel...")
        try:
            # Import and run the    converter directly
            import json_to_xlsx
            json_to_xlsx.create_excel_from_json(
                data=json_to_xlsx.load_json_result('ai_responses/schedule_result.json'),
                output_file='output_transformed.xlsx'
            )
            print("  ✓ Excel conversion completed!")
        except Exception as e:
            print(f"  ✗ Error during Excel conversion: {e}")
            print("  💡 Try running manually: python json_to_xlsx.py")

        print("\n" + "=" * 60)
        print("All tasks completed!")
        print("=" * 60)

    except Exception as e:
        print(f"\nProgram execution failed: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n[Interrupted] Operation cancelled by user")
    except Exception as e:
        print(f"\n[Fatal Error] {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("\n" + "=" * 60)
        print("Program ended")
        print("=" * 60)