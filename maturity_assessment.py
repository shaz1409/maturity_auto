#!/usr/bin/env python3
"""
Marketing Maturity Assessment Automation

Automated script to:
1. Load survey data from Google Sheets
2. Calculate category scores for each client
3. Generate AI-powered recommendations
4. Create personalized PowerPoint presentations

Usage:
    python maturity_assessment.py

Environment Variables:
    OPENAI_API_KEY: Your OpenAI API key (required)
"""

import pandas as pd
import numpy as np
import requests
import io
import urllib3
import os
import re
import sys
from datetime import datetime
from pptx import Presentation

# Suppress SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configuration
SHEET_ID = '1tHWUJWJl_zTwGRTg21qbW_5qpYH8bBMFYZYbIZ03eO8'
GID = '491555971'
SHEET_URL = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}'
TEMPLATE_PATH = 'Maturity_Slide_Template.pptx'
OUTPUT_DIR = 'output'

# Initialize OpenAI client
try:
    from openai import OpenAI
    api_key = os.environ.get('OPENAI_API_KEY')
    if not api_key or api_key == 'your-api-key-here':
        print("ERROR: OPENAI_API_KEY not set. Please set it as an environment variable.")
        sys.exit(1)
    client = OpenAI(api_key=api_key)
except ImportError:
    print("ERROR: OpenAI package not installed. Run: pip install openai")
    sys.exit(1)
except Exception as e:
    print(f"ERROR: Failed to initialize OpenAI client: {e}")
    sys.exit(1)


def clean_column_name(col_name):
    """Clean column name: lowercase, remove ALL special characters, replace spaces with underscores"""
    cleaned = col_name.lower()
    cleaned = cleaned.replace('-', ' ')
    cleaned = re.sub(r'[^a-z0-9\s]', '', cleaned)
    cleaned = re.sub(r'\s+', ' ', cleaned)
    cleaned = cleaned.strip().replace(' ', '_')
    return cleaned


def load_data():
    """Load survey data from Google Sheets"""
    print("Loading data from Google Sheet...")
    response = requests.get(SHEET_URL, verify=False, timeout=30)
    response.raise_for_status()
    df = pd.read_csv(io.StringIO(response.text))
    print(f"✓ Data loaded: {df.shape[0]} rows, {df.shape[1]} columns")
    return df


def setup_mappings(df):
    """Set up column name mappings and question categories"""
    # Create column name mappings
    COLUMN_NAME_MAPPING = {}
    CLEANED_TO_ORIGINAL = {}
    
    for col in df.columns:
        cleaned = clean_column_name(col)
        COLUMN_NAME_MAPPING[col] = cleaned
        CLEANED_TO_ORIGINAL[cleaned] = col
    
    # Get question columns in order
    question_columns = [col for col in df.columns if col not in ['Timestamp', 'Email Address']]
    cleaned_questions = [COLUMN_NAME_MAPPING[q] for q in question_columns]
    
    # Define question categories
    QUESTION_CATEGORIES = {
        'Tech & Data': cleaned_questions[0:5],
        'Campaigning & Assets': cleaned_questions[5:11],
        'Segmentation & Personalisation': cleaned_questions[11:14],
        'Reporting & Insights': cleaned_questions[14:20],
        'People & Operations': cleaned_questions[20:24]
    }
    
    # Create reverse mapping
    QUESTION_TO_CATEGORY = {}
    for category, questions in QUESTION_CATEGORIES.items():
        for question in questions:
            QUESTION_TO_CATEGORY[question] = category
    
    # Create cleaned to original column mapping
    CLEANED_TO_ORIGINAL_COL = {cleaned: original for original, cleaned in COLUMN_NAME_MAPPING.items()}
    
    return (COLUMN_NAME_MAPPING, CLEANED_TO_ORIGINAL_COL, QUESTION_CATEGORIES, 
            QUESTION_TO_CATEGORY, question_columns)


def calculate_category_scores(client_row, question_categories, cleaned_to_original):
    """Calculate average score for each category for a single client"""
    scores = {}
    
    for category, questions in question_categories.items():
        category_scores = []
        
        for cleaned_q in questions:
            original_col = cleaned_to_original.get(cleaned_q)
            
            if original_col:
                value = client_row.get(original_col, None)
                
                if value is not None and pd.notna(value):
                    try:
                        num_value = float(value)
                        if 1 <= num_value <= 4:
                            category_scores.append(num_value)
                    except (ValueError, TypeError):
                        pass
        
        if category_scores:
            scores[category] = np.mean(category_scores)
        else:
            scores[category] = None
    
    return scores


def find_text_boxes(slide):
    """Find text boxes and orange circle on a slide"""
    elements = {
        'score': None,
        'recommendations': None,
        'orange_circle': None,
        'line': None
    }
    
    # Find text boxes
    for shape in slide.shapes:
        if hasattr(shape, "text"):
            text = shape.text.strip()
            if "Your score" in text and elements['score'] is None:
                elements['score'] = shape
            elif "Recommendation 1" in text and elements['recommendations'] is None:
                elements['recommendations'] = shape
    
    # Find line and circle
    lines = []
    circles = []
    
    for shape in slide.shapes:
        if hasattr(shape, 'shape_type'):
            if shape.shape_type == 9:  # LINE
                lines.append(shape)
            elif shape.shape_type == 1:  # AUTO_SHAPE
                if hasattr(shape, 'width') and hasattr(shape, 'height'):
                    width = shape.width
                    height = shape.height
                    if abs(width - height) / max(width, height) < 0.2:
                        try:
                            if hasattr(shape, 'fill') and shape.fill.type == 1:
                                circles.append(shape)
                        except:
                            pass
    
    if lines:
        elements['line'] = max(lines, key=lambda l: l.width if hasattr(l, 'width') else 0)
    
    if circles and elements['line']:
        line_y = elements['line'].top + (elements['line'].height / 2)
        elements['orange_circle'] = min(circles, 
                                        key=lambda c: abs((c.top + c.height/2) - line_y))
    elif circles:
        elements['orange_circle'] = circles[0]
    
    return elements


def generate_recommendations(category, score, questions_in_category, client_responses, original_questions_dict):
    """Generate recommendations using OpenAI"""
    # Identify low-scoring questions
    low_scoring_questions = []
    all_questions_context = []
    
    for cleaned_q in questions_in_category:
        q_score = client_responses.get(cleaned_q, None)
        if q_score is not None and isinstance(q_score, (int, float)):
            original_q = original_questions_dict.get(cleaned_q, cleaned_q)
            question_info = {
                'cleaned': cleaned_q,
                'original': original_q,
                'score': q_score
            }
            all_questions_context.append(question_info)
            
            if q_score <= 2:
                low_scoring_questions.append(question_info)
    
    low_scoring_questions.sort(key=lambda x: x['score'])
    all_questions_context.sort(key=lambda x: x['score'])
    
    questions_detail = "\n".join([
        f"- Question: {q['original']}\n  Score: {q['score']}/4"
        for q in all_questions_context
    ])
    
    focus_areas = ""
    if low_scoring_questions:
        focus_areas = "\n\nAreas requiring immediate attention (low scores):\n" + "\n".join([
            f"- {q['original']} (Score: {q['score']}/4)"
            for q in low_scoring_questions[:3]
        ])
    
    # Determine maturity level
    if score <= 1.5:
        maturity_level = "not mature"
    elif score <= 2.5:
        maturity_level = "developing"
    elif score <= 3.5:
        maturity_level = "mature"
    else:
        maturity_level = "very mature"
    
    prompt = f"""You are a CRM marketing maturity consultant. Generate recommendations for a client.

Category: {category}
Overall Maturity Score: {score:.2f}/4.0 ({maturity_level})

Client's responses to all questions in this category:
{questions_detail}
{focus_areas}

Instructions:
1. Generate a brief 2-3 sentence summary of their current maturity level in this category
2. Generate four specific, actionable recommendations that:
   - PRIORITIZE addressing the low-scoring areas identified above
   - Reference the specific questions where they scored low (1-2 out of 4)
   - Provide concrete, actionable steps based on the question context
   - Are tailored to their current maturity level

Focus especially on the questions where they scored 1-2, as these are the areas needing the most improvement.

Format the response as:
SUMMARY: [your summary here]
RECOMMENDATIONS:
1. [recommendation 1 - should address a specific low-scoring question]
2. [recommendation 2 - should address a specific low-scoring question]
3. [recommendation 3 - can address another area or build on improvements]
4. [recommendation 4 - can address another area or build on improvements]

Make each recommendation specific, actionable, and directly related to the questions they answered poorly."""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are an expert CRM marketing consultant providing actionable recommendations."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=500
        )
        
        result = response.choices[0].message.content
        
        if "SUMMARY:" in result:
            parts = result.split("RECOMMENDATIONS:")
            summary = parts[0].replace("SUMMARY:", "").strip()
            recommendations_text = parts[1] if len(parts) > 1 else ""
            
            recommendations = []
            for line in recommendations_text.split('\n'):
                line = line.strip()
                if line and (line[0].isdigit() or line.startswith('-')):
                    rec = line.split('.', 1)[-1].strip() if '.' in line else line.lstrip('- ').strip()
                    if rec:
                        recommendations.append(rec)
            
            while len(recommendations) < 4:
                recommendations.append("Continue building on the recommendations above.")
            
            return summary[:200], recommendations[:4]
        else:
            lines = result.split('\n')
            summary = lines[0][:200] if lines else "Summary generated"
            recommendations = [line.strip() for line in lines[1:] if line.strip() and len(line.strip()) > 10][:4]
            while len(recommendations) < 4:
                recommendations.append("Continue improving in this area.")
            return summary, recommendations
            
    except Exception as e:
        print(f"  ⚠️  Error generating recommendations: {e}")
        return f"Error: {str(e)}", [
            "Recommendation 1: Review current processes",
            "Recommendation 2: Identify improvement areas",
            "Recommendation 3: Implement best practices",
            "Recommendation 4: Monitor progress"
        ]


def map_slides_to_categories(prs):
    """Map slide titles to category names"""
    SLIDE_CATEGORY_MAPPING = {
        'Tech and Data': 'Tech & Data',
        'Campaigning & Assets': 'Campaigning & Assets',
        'Segmentation & Personalisation': 'Segmentation & Personalisation',
        'Reporting & Insights': 'Reporting & Insights',
        'People & Operations': 'People & Operations'
    }
    
    CATEGORY_TO_SLIDE = {}
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if hasattr(shape, 'is_placeholder') and shape.is_placeholder:
                title = shape.text.strip()
                if title in SLIDE_CATEGORY_MAPPING:
                    category = SLIDE_CATEGORY_MAPPING[title]
                    CATEGORY_TO_SLIDE[category] = i
                    break
    
    return CATEGORY_TO_SLIDE


def email_to_filename(email):
    """Convert email to safe filename"""
    safe_email = email.replace('@', '_at_').replace('.', '_')
    return f"{safe_email}_Maturity_Assessment.pptx"


def generate_client_presentation(client_email, client_scores, client_responses, 
                                  COLUMN_NAME_MAPPING, CLEANED_TO_ORIGINAL_COL,
                                  QUESTION_CATEGORIES, output_filename):
    """Generate PowerPoint presentation for a client"""
    prs = Presentation(TEMPLATE_PATH)
    CATEGORY_TO_SLIDE = map_slides_to_categories(prs)
    
    # Map responses
    question_responses = {}
    for orig_col, cleaned_col in COLUMN_NAME_MAPPING.items():
        if orig_col in client_responses:
            question_responses[cleaned_col] = client_responses[orig_col]
    
    # Process each category
    for category, score in client_scores.items():
        if category not in CATEGORY_TO_SLIDE:
            continue
        
        if pd.isna(score) or score is None:
            continue
        
        slide_idx = CATEGORY_TO_SLIDE[category]
        slide = prs.slides[slide_idx]
        elements = find_text_boxes(slide)
        
        questions_in_category = QUESTION_CATEGORIES.get(category, [])
        category_responses = {q: question_responses.get(q, 'N/A') for q in questions_in_category}
        
        original_questions_dict = {}
        for cleaned_q in questions_in_category:
            orig_col = CLEANED_TO_ORIGINAL_COL.get(cleaned_q)
            if orig_col:
                original_questions_dict[cleaned_q] = orig_col
        
        # Generate recommendations
        print(f"  Generating recommendations for {category} (score: {score:.2f})...")
        summary, recommendations = generate_recommendations(
            category, score, questions_in_category, category_responses, original_questions_dict
        )
        
        # Update score
        if elements['score']:
            if hasattr(elements['score'], 'text_frame'):
                elements['score'].text_frame.clear()
                p = elements['score'].text_frame.paragraphs[0]
                p.text = f"{score:.2f}/4.0"
            else:
                elements['score'].text = f"{score:.2f}/4.0"
        
        # Update recommendations
        if elements['recommendations']:
            recommendations_text = "\n\n".join([f"{i+1}. {rec}" for i, rec in enumerate(recommendations)])
            if hasattr(elements['recommendations'], 'text_frame'):
                elements['recommendations'].text_frame.clear()
                p = elements['recommendations'].text_frame.paragraphs[0]
                p.text = recommendations_text
            else:
                elements['recommendations'].text = recommendations_text
        
        # Position orange circle
        if elements['orange_circle'] and elements['line']:
            try:
                line = elements['line']
                line_left = line.left
                line_width = line.width
                line_top = line.top
                line_height = line.height
                
                score_position = score / 4.0
                circle_x = line_left + (line_width * score_position)
                circle_y = line_top + (line_height / 2)
                
                circle = elements['orange_circle']
                circle_width = circle.width
                circle_height = circle.height
                
                circle.left = int(circle_x - (circle_width / 2))
                circle.top = int(circle_y - (circle_height / 2))
            except Exception as e:
                print(f"    ⚠️  Could not position orange circle: {e}")
    
    prs.save(output_filename)
    return output_filename


def main():
    """Main execution function"""
    print("="*60)
    print("Marketing Maturity Assessment Automation")
    print(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)
    
    # Create output directory
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    try:
        # Load data
        df = load_data()
        
        # Setup mappings
        print("\nSetting up mappings...")
        (COLUMN_NAME_MAPPING, CLEANED_TO_ORIGINAL_COL, QUESTION_CATEGORIES,
         QUESTION_TO_CATEGORY, question_columns) = setup_mappings(df)
        print(f"✓ Mapped {len(QUESTION_CATEGORIES)} categories")
        
        # Calculate scores for all clients
        print("\nCalculating category scores...")
        category_scores_list = []
        
        for idx, row in df.iterrows():
            client_scores = calculate_category_scores(row, QUESTION_CATEGORIES, CLEANED_TO_ORIGINAL_COL)
            client_scores['Email Address'] = row['Email Address']
            client_scores['Timestamp'] = row['Timestamp']
            category_scores_list.append(client_scores)
        
        scores_df = pd.DataFrame(category_scores_list)
        print(f"✓ Calculated scores for {len(scores_df)} clients")
        
        # Generate presentations
        print(f"\nGenerating presentations...")
        generated_files = []
        
        for idx, row in scores_df.iterrows():
            client_email = row['Email Address']
            client_scores = {cat: row[cat] for cat in QUESTION_CATEGORIES.keys()}
            
            client_row = df[df['Email Address'] == client_email].iloc[0]
            client_responses = {}
            for col in question_columns:
                client_responses[col] = client_row[col]
            
            filename = email_to_filename(client_email)
            output_filename = os.path.join(OUTPUT_DIR, filename)
            
            print(f"\n  Processing: {client_email}")
            try:
                generate_client_presentation(
                    client_email, client_scores, client_responses,
                    COLUMN_NAME_MAPPING, CLEANED_TO_ORIGINAL_COL,
                    QUESTION_CATEGORIES, output_filename
                )
                generated_files.append(output_filename)
                print(f"    ✓ Saved: {filename}")
            except Exception as e:
                print(f"    ❌ Error: {e}")
        
        print(f"\n{'='*60}")
        print(f"✓ Successfully generated {len(generated_files)} presentations")
        print(f"Saved to: {OUTPUT_DIR}/")
        print(f"Completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("="*60)
        
    except Exception as e:
        print(f"\n❌ Fatal error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
