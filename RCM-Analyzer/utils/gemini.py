import os
import google.generativeai as genai
from typing import Dict, List, Any, Union
import logging
import json

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def initialize_gemini():
    """
    Initialize Gemini API client with the API key
    
    Returns:
        Gemini API model instance
    """
    try:
        # Get API key from environment variable
        api_key = os.environ.get("GEMINI_API_KEY")
        
        if not api_key:
            # Check if .env file has been loaded but variable name is different
            for env_var in os.environ.keys():
                if "GEMINI" in env_var.upper() and "KEY" in env_var.upper() and os.environ.get(env_var):
                    api_key = os.environ.get(env_var)
                    logger.info(f"Using API key from {env_var}")
                    break
            
            if not api_key:
                raise ValueError("GEMINI_API_KEY environment variable not found. Please set it.")
        
        # Configure the Gemini API
        genai.configure(api_key=api_key)
        
        # Directly use Gemini 1.5 Flash model
        model_name = "models/gemini-1.5-flash-latest"
        logger.info(f"Using Gemini 1.5 Flash model: {model_name}")
        
        # Initialize the model
        model = genai.GenerativeModel(model_name)
        
        # Set generation config
        model.generation_config = {
            "temperature": 0.2,
            "top_p": 0.9,
            "top_k": 32,
            "max_output_tokens": 8192,
        }
        
        return model
    
    except Exception as e:
        logger.error(f"Error initializing Gemini: {str(e)}")
        raise

def analyze_risk_with_gemini(model, data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Analyze RCM data with Gemini, focusing on departmental risks
    
    Args:
        model: Gemini model instance
        data: Structured data from document processing
        
    Returns:
        Enhanced data with Gemini's analysis
    """
    try:
        logger.info("Starting Risk Control Matrix analysis with Gemini")
        enhanced_data = data.copy()
        
        # First, check if we have raw data to use for RAG
        if "raw_data" in data and data["raw_data"]:
            logger.info(f"Using RAG approach with {len(data['raw_data'])} sheets of raw data")
            return analyze_with_rag(model, data)
        # If this is raw text (from PDF or DOCX), we need to perform structured extraction
        elif "raw_text" in data and data["raw_text"]:
            logger.info("Processing raw text document")
            return analyze_raw_document(model, data)
        # For structured data without raw_data, enhance it with departmental risk analysis
        else:
            logger.info("Using standard analysis for structured data")
            # Check if we need to generate department_risks structures
            if not data.get("department_risks"):
                logger.info("No department_risks found in processed data, generating...")
                enhanced_data["department_risks"] = generate_department_risk_matrix(data)
            
            # Transform department_risks from just risk categories to full analysis
            department_risks = enhanced_data.get("department_risks", {})
            enhanced_dept_risks = {}
            
            for dept, risk_data in department_risks.items():
                # Check if it's just risk categories or already a full analysis
                if isinstance(risk_data, dict) and all(key in risk_data for key in ["overall_risk_level", "risk_categories", "key_risks", "summary"]):
                    # Already a full analysis
                    enhanced_dept_risks[dept] = risk_data
                else:
                    # Just risk categories, generate full analysis
                    risk_categories = risk_data if isinstance(risk_data, dict) else {}
                    if risk_categories:
                        # Calculate overall risk level based on category values
                        category_values = list(risk_categories.values())
                        avg_risk = sum(category_values) / len(category_values) if category_values else 0
                        
                        if avg_risk >= 3.5:
                            overall_risk = "High"
                        elif avg_risk >= 2.5:
                            overall_risk = "Medium"
                        else:
                            overall_risk = "Low"
                        
                        # Get department-specific objectives
                        dept_objectives = [obj for obj in data.get("control_objectives", []) if obj.get("department") == dept]
                        
                        # Use Gemini to analyze this department
                        dept_analysis = analyze_department(model, dept, dept_objectives, risk_categories)
                        enhanced_dept_risks[dept] = dept_analysis
            
            enhanced_data["department_risks"] = enhanced_dept_risks
            
            # Generate risk score if not present
            if "risk_score" not in enhanced_data:
                enhanced_data["risk_score"] = calculate_risk_score(enhanced_data)
            
            # Generate recommendations if not present
            if "recommendations" not in enhanced_data:
                enhanced_data["recommendations"] = generate_department_recommendations(model, enhanced_data)
            
            return enhanced_data
    
    except Exception as e:
        logger.error(f"Error analyzing data with Gemini: {str(e)}")
        raise

def analyze_with_rag(model, data: Dict[str, Any]) -> Dict[str, Any]:
    """
    True RAG approach - send raw data directly to Gemini for comprehensive analysis
    """
    enhanced_data = data.copy()
    
    # Format raw data into a structured format for the LLM
    raw_data_text = format_raw_data(data["raw_data"])
    
    # Get list of departments to analyze
    departments = data.get("departments", [])
    departments_str = ", ".join(departments) if departments else "All departments identified in the Area column"
    
    # Prepare the prompt for the LLM to analyze the raw data
    prompt = f"""
    You are a Risk Control Matrix (RCM) analysis expert. I will provide you with data from an uploaded RCM file.
    
    Your task is to analyze all departments in the file, especially focusing on these specific departments:
    {departments_str}
    
    For each department, you must:
    1. Analyze the control objectives and risks
    2. Classify risks into these categories: Operational, Financial, Fraud, Financial Fraud, Operational Fraud
    3. Identify control gaps
    4. Provide recommendations

    In the RCM file, departments are shown in the "Area" column, which includes values like "Employee Master Maintenance", "Attendance & Payroll Processing", etc.
    
    Here is the raw data from the RCM document:
    
    {raw_data_text}
    
    Analyze ALL departments found in the data. For each department, provide a comprehensive analysis in this JSON format:
    
    {{
        "departments": [
            {{
                "name": "Department Name",
                "overall_risk_level": "High/Medium/Low",
                "key_risks": ["Risk 1", "Risk 2", ...],
                "risk_analysis": {{
                    "Operational": ["Specific operational risks"],
                    "Financial": ["Specific financial risks"],
                    "Fraud": ["Specific fraud risks"],
                    "Financial Fraud": ["Specific financial fraud risks"],
                    "Operational Fraud": ["Specific operational fraud risks"]
                }},
                "control_gaps": [
                    {{
                        "gap_title": "Gap description",
                        "impact": "Impact description",
                        "recommendation": "Recommendation to address gap"
                    }}
                ],
                "summary": "Brief summary of department's risk profile"
            }}
        ],
        "overall_recommendations": [
            {{
                "title": "Recommendation title",
                "priority": "High/Medium/Low",
                "description": "Detailed recommendation",
                "impact": "Expected impact of implementation"
            }}
        ]
    }}
    
    IMPORTANT: You MUST analyze ALL departments found in the RCM data, especially those in the Area column like "Employee Master Maintenance", "Attendance & Payroll Processing", "Payroll and Personnel", "Leave Management", and "Separation". Do not focus on only one department.
    
    Be comprehensive, but focus on practical, actionable insights. Identify specific risks rather than general statements.
    """
    
    # Get response from Gemini
    logger.info("Sending RAG prompt to Gemini")
    response = model.generate_content(prompt)
    
    # Extract JSON from the response
    try:
        response_text = response.text
        # Check if response has JSON code blocks and extract them
        if "```json" in response_text:
            json_text = response_text.split("```json")[1].split("```")[0].strip()
        elif "```" in response_text:
            json_text = response_text.split("```")[1].strip()
        else:
            json_text = response_text.strip()
        
        rag_analysis = json.loads(json_text)
        logger.info(f"Successfully parsed RAG analysis results: {len(rag_analysis.get('departments', []))} departments")
        
        # Transform the RAG analysis into our structured format
        enhanced_data["department_risks"] = {}
        enhanced_data["recommendations"] = []
        
        # Process departments
        for dept_data in rag_analysis.get("departments", []):
            dept_name = dept_data.get("name", "Unknown")
            
            # Create department risk entry
            enhanced_data["department_risks"][dept_name] = {
                "overall_risk_level": dept_data.get("overall_risk_level", "Medium"),
                "key_risks": dept_data.get("key_risks", []),
                "risk_types": dept_data.get("risk_analysis", {}),
                "summary": dept_data.get("summary", ""),
                "risk_categories": {
                    "Financial": 4 if dept_data.get("risk_analysis", {}).get("Financial", []) else 2,
                    "Operational": 4 if dept_data.get("risk_analysis", {}).get("Operational", []) else 2,
                    "Compliance": 3,  # Default
                    "Strategic": 3,   # Default
                    "Technological": 3  # Default
                }
            }
            
            # Add gaps
            for gap in dept_data.get("control_gaps", []):
                enhanced_data["gaps"].append({
                    "department": dept_name,
                    "gap_title": gap.get("gap_title", ""),
                    "description": gap.get("gap_title", ""),
                    "risk_impact": gap.get("impact", ""),
                    "proposed_solution": gap.get("recommendation", "")
                })
        
        # Add overall recommendations
        for rec in rag_analysis.get("overall_recommendations", []):
            enhanced_data["recommendations"].append({
                "title": rec.get("title", ""),
                "priority": rec.get("priority", "Medium"),
                "description": rec.get("description", ""),
                "impact": rec.get("impact", ""),
                "complexity": "Medium"  # Default
            })
        
        # If we don't have all the departments, add defaults for missing ones
        if departments and len(enhanced_data["department_risks"]) < len(departments):
            missing_depts = [d for d in departments if d not in enhanced_data["department_risks"]]
            for dept in missing_depts:
                logger.warning(f"Department {dept} was missing from analysis, adding default")
                enhanced_data["department_risks"][dept] = {
                    "overall_risk_level": "Medium",
                    "key_risks": [f"Need to analyze {dept} department risks"],
                    "risk_types": {
                        "Operational": [f"Potential operational risks in {dept}"],
                        "Financial": [],
                        "Fraud": [],
                        "Financial Fraud": [],
                        "Operational Fraud": []
                    },
                    "summary": f"Additional analysis required for {dept} department",
                    "risk_categories": {
                        "Financial": 2,
                        "Operational": 3,
                        "Compliance": 2,
                        "Strategic": 2,
                        "Technological": 2
                    }
                }
        
        return enhanced_data
        
    except Exception as json_error:
        logger.error(f"Error extracting JSON from Gemini RAG response: {str(json_error)}")
        logger.debug(f"Raw response: {response.text}")
        # Fall back to standard analysis
        return analyze_structured_data(model, data)

def format_raw_data(raw_data):
    """Format raw data from Excel sheets into a structured text representation"""
    formatted_text = ""
    
    for sheet_data in raw_data:
        sheet_name = sheet_data.get("sheet_name", "Unknown Sheet")
        rows = sheet_data.get("rows", [])
        
        formatted_text += f"\n=== SHEET: {sheet_name} ===\n\n"
        
        # If there are rows, format them
        if rows:
            # Get columns from first row
            if rows:
                columns = list(rows[0].keys())
                formatted_text += "| " + " | ".join(columns) + " |\n"
                formatted_text += "| " + " | ".join(["-" * len(col) for col in columns]) + " |\n"
                
                # Add rows
                for row in rows:
                    formatted_text += "| " + " | ".join([str(row.get(col, "")) for col in columns]) + " |\n"
        
        formatted_text += "\n"
    
    return formatted_text

def analyze_structured_data(model, data: Dict[str, Any]) -> Dict[str, Any]:
    """Analyze already structured data (fallback if RAG fails)"""
    enhanced_data = data.copy()
    
    # Generate department risk matrix if not present
    if not data.get("department_risks"):
        department_risks = generate_department_risk_matrix(data)
        enhanced_data["department_risks"] = department_risks
    
    # Generate recommendations
    enhanced_data["recommendations"] = generate_department_recommendations(model, data)
    
    # Calculate risk score
    enhanced_data["risk_score"] = calculate_risk_score(data)
    
    return enhanced_data

def analyze_raw_document(model, data: Dict[str, Any]) -> Dict[str, Any]:
    """Process raw text documents using Gemini"""
    enhanced_data = data.copy()
    
    # Prepare the prompt for Gemini
    if "extracted_text" in data:
        extracted_text = data["extracted_text"]
        
        # Truncate if too long and warn
        if len(extracted_text) > 30000:
            logger.warning(f"Extracted text is very long ({len(extracted_text)} chars). Truncating for Gemini analysis.")
            extracted_text = extracted_text[:30000] + "..."
        
        prompt = f"""
        You are a Risk Assessment and Control expert. I will provide you with text from a Risk Control Matrix (RCM) document.
        
        Please analyze this text and extract the following structured information, with special focus on departmental risks:
        
        1. Departments: Identify all departments mentioned in the document.
        2. Control Objectives: For each department, identify the main control objectives.
        3. What Can Go Wrong: For each control objective, identify what could go wrong if the control is not implemented.
        4. Risk Levels: Identify the risk level (High, Medium, Low) for each control objective.
        5. Control Activities: Identify the control activities in place to address each risk.
        6. Gaps: Identify any control or design gaps mentioned in the document.
        7. Proposed Controls: Identify any proposed controls to address the gaps.
        8. Departmental Risk Analysis: Provide a risk assessment for each department, including risk categories and overall risk level.
        
        Please be comprehensive and detailed in your analysis. Here is the text:
        
        {extracted_text}
        
        Respond with ONLY a JSON object containing the extracted structured information. The format should be:
        {{
            "departments": ["string"],
            "control_objectives": [
                {{
                    "department": "string",
                    "objective": "string",
                    "what_can_go_wrong": "string",
                    "risk_level": "string",
                    "control_activities": "string",
                    "is_gap": boolean,
                    "gap_details": "string",
                    "proposed_control": "string"
                }}
            ],
            "gaps": [
                {{
                    "department": "string",
                    "control_objective": "string",
                    "gap_title": "string",
                    "description": "string",
                    "risk_impact": "string",
                    "proposed_solution": "string"
                }}
            ],
            "department_risks": {{
                "Department1": {{
                    "overall_risk_level": "string",
                    "risk_categories": {{
                        "Financial": number,
                        "Operational": number,
                        "Compliance": number,
                        "Strategic": number,
                        "Technological": number
                    }},
                    "key_risks": ["string"],
                    "summary": "string"
                }}
            }},
            "risk_distribution": {{"Low": number, "Medium": number, "High": number}},
            "total_controls": number,
            "control_gaps": number
        }}
        """
        
        # Get response from Gemini
        response = model.generate_content(prompt)
        
        # Extract JSON from the response
        try:
            response_text = response.text
            # Check if response has JSON code blocks and extract them
            if "```json" in response_text:
                json_text = response_text.split("```json")[1].split("```")[0].strip()
            elif "```" in response_text:
                json_text = response_text.split("```")[1].strip()
            else:
                json_text = response_text.strip()
            
            extracted_data = json.loads(json_text)
            
            # Update the enhanced data with extracted information
            enhanced_data.update(extracted_data)
            enhanced_data["raw_text"] = False  # Mark as processed
            
            # Additional post-processing
            enhanced_data["risk_score"] = calculate_risk_score(extracted_data)
            enhanced_data["recommendations"] = generate_department_recommendations(model, extracted_data)
            
        except Exception as json_error:
            logger.error(f"Error extracting JSON from Gemini response: {str(json_error)}")
            logger.debug(f"Raw response: {response.text}")
            raise
    
    return enhanced_data

def generate_department_risk_matrix(data: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    """Generate a basic department risk matrix from structured data"""
    departments = data.get("departments", [])
    if not departments:
        return {}
    
    risk_categories = ["Financial", "Operational", "Compliance", "Strategic", "Technological"]
    department_risks = {dept: {cat: 0 for cat in risk_categories} for dept in departments}
    
    # Populate risk values based on control objectives
    for obj in data.get("control_objectives", []):
        dept = obj.get("department", "")
        if not dept or dept not in departments:
            continue
            
        # Set risk level value
        risk_text = obj.get("risk_level", "").lower()
        if risk_text in ['high', 'h', 'critical', 'severe']:
            risk_level_value = 4
        elif risk_text in ['medium', 'm', 'mod', 'moderate']:
            risk_level_value = 3
        else:
            risk_level_value = 2
            
        # Assign to categories based on keywords
        objective_text = obj.get("objective", "").lower()
        what_can_go_wrong = obj.get("what_can_go_wrong", "").lower()
        combined_text = f"{objective_text} {what_can_go_wrong}"
        
        # Financial risks
        if any(term in combined_text for term in ['financ', 'account', 'budget', 'cost', 'expense', 'revenue', 'payment', 'tax', 'audit']):
            department_risks[dept]["Financial"] = max(department_risks[dept]["Financial"], risk_level_value)
            
        # Operational risks
        if any(term in combined_text for term in ['operat', 'process', 'procedur', 'workflow', 'efficien', 'product', 'service', 'delivery']):
            department_risks[dept]["Operational"] = max(department_risks[dept]["Operational"], risk_level_value)
            
        # Compliance risks
        if any(term in combined_text for term in ['comply', 'compliance', 'regulat', 'legal', 'law', 'policy', 'requirement', 'standard']):
            department_risks[dept]["Compliance"] = max(department_risks[dept]["Compliance"], risk_level_value)
            
        # Strategic risks
        if any(term in combined_text for term in ['strateg', 'goal', 'objective', 'mission', 'vision', 'plan', 'market', 'competi']):
            department_risks[dept]["Strategic"] = max(department_risks[dept]["Strategic"], risk_level_value)
            
        # Technological risks
        if any(term in combined_text for term in ['tech', 'system', 'data', 'secur', 'access', 'software', 'hardware', 'it ', 'cyber']):
            department_risks[dept]["Technological"] = max(department_risks[dept]["Technological"], risk_level_value)
    
    # Ensure all departments have at least some risk level for each category
    for dept in departments:
        for cat in risk_categories:
            if department_risks[dept][cat] == 0:
                department_risks[dept][cat] = 1  # Set minimum risk level
    
    return department_risks

def analyze_department(model, dept: str, objectives: List[Dict[str, Any]], risk_categories: Dict[str, int]) -> Dict[str, Any]:
    """Use Gemini to analyze a specific department"""
    try:
        # Format objectives for the prompt
        objectives_text = ""
        for i, obj in enumerate(objectives[:5]):  # Limit to 5 for brevity
            objectives_text += f"{i+1}. Objective: {obj.get('objective', '')}\n"
            objectives_text += f"   What Can Go Wrong: {obj.get('what_can_go_wrong', '')}\n"
            objectives_text += f"   Risk Level: {obj.get('risk_level', '')}\n"
            objectives_text += f"   Control Activities: {obj.get('control_activities', '')}\n"
            if obj.get("is_gap", False):
                objectives_text += f"   Gap: {obj.get('gap_details', '')}\n"
            objectives_text += "\n"
            
        # Add risk categories
        categories_text = "\nRisk Categories:\n"
        for cat, value in risk_categories.items():
            categories_text += f"- {cat}: {value}/5\n"
        
        # Calculate overall risk level based on category values
        category_values = list(risk_categories.values())
        avg_risk = sum(category_values) / len(category_values) if category_values else 0
        
        if avg_risk >= 3.5:
            suggested_risk = "High"
        elif avg_risk >= 2.5:
            suggested_risk = "Medium"
        else:
            suggested_risk = "Low"
        
        # Create a dedicated solutions prompt for each objective
        solutions_prompts = []
        for i, obj in enumerate(objectives):
            what_can_go_wrong = obj.get("what_can_go_wrong", "")
            objective = obj.get("objective", "")
            risk_level = obj.get("risk_level", "Medium")
            
            solutions_prompts.append(f"""
            Objective {i+1}: {objective}
            What Can Go Wrong: {what_can_go_wrong}
            Risk Level: {risk_level}
            
            For this specific control objective, provide a detailed, tailored solution that directly addresses the risk.
            """)
        
        combined_solutions_prompt = "\n\n".join(solutions_prompts)
        
        prompt = f"""
        You are a Risk Management and Internal Controls Expert with extensive experience in designing control frameworks and providing solutions to address control gaps. I need your help to analyze control objectives for the {dept} department and provide specific, actionable solutions.

        DEPARTMENT: {dept}
        
        RISK CATEGORIES:
        {categories_text}
        
        CONTROL OBJECTIVES:
        {objectives_text}
        
        TASK:
        For EACH control objective above, you must provide:
        1. A determination of whether there is a control design gap (Yes/No)
        2. A UNIQUE, DETAILED proposed solution (approximately 50 words, 2-3 sentences) that specifically addresses the risk described in "What Can Go Wrong"
        
        REQUIREMENTS FOR PROPOSED SOLUTIONS:
        - Each solution must be tailored to the specific control objective and risk
        - Solutions must be practical, actionable, and implementable
        - Solutions must include specific technologies, processes, or controls to implement
        - AVOID generic solutions that could apply to any risk
        - AVOID reusing the same or similar solutions for multiple objectives
        - Each solution should be approximately 50 words (2-3 detailed sentences)
        
        I NEED SPECIFIC SOLUTIONS FOR EACH OF THESE CONTROL OBJECTIVES:
        {combined_solutions_prompt}
        
        Respond with ONLY a JSON object in the following format:
        {{
            "overall_risk_level": "High/Medium/Low",
            "risk_categories": {json.dumps(risk_categories)},
            "key_risks": ["3-5 key risks identified"],
            "risk_types": {{
                "Operational": ["specific operational risks"],
                "Financial": ["specific financial risks"],
                "Fraud": ["specific fraud risks"],
                "Financial_Fraud": ["specific financial fraud risks"],
                "Operational_Fraud": ["specific operational fraud risks"]
            }},
            "summary": "brief department risk summary",
            "control_gaps": [
                {{
                    "objective": "exact text of the control objective",
                    "has_gap": "Yes/No",
                    "proposed_solution": "unique, tailored solution of approximately 50 words"
                }}
            ]
        }}
        """
        
        # Get response from Gemini
        response = model.generate_content(prompt)
        
        # Extract JSON
        response_text = response.text
        if "```json" in response_text:
            json_text = response_text.split("```json")[1].split("```")[0].strip()
        elif "```" in response_text:
            json_text = response_text.split("```")[1].strip()
        else:
            json_text = response_text.strip()
        
        dept_analysis = json.loads(json_text)
        
        # Process control gaps and update objectives with gap information
        if "control_gaps" in dept_analysis:
            for gap_info in dept_analysis["control_gaps"]:
                obj_text = gap_info.get("objective", "")
                has_gap = gap_info.get("has_gap", "No")
                proposed_solution = gap_info.get("proposed_solution", "")
                
                # Find matching objective and update it
                for obj in objectives:
                    if obj.get("objective", "") == obj_text or obj_text in obj.get("objective", ""):
                        obj["gap_details"] = "Yes" if has_gap.lower() == "yes" else "No"
                        # Always add the proposed solution regardless of gap status
                        if proposed_solution:
                            obj["proposed_control"] = proposed_solution
        
        return dept_analysis
        
    except Exception as e:
        logger.error(f"Error analyzing department {dept}: {str(e)}")
        # Create fallback analysis
        return {
            "overall_risk_level": suggested_risk,
            "risk_categories": risk_categories,
            "key_risks": [
                f"{dept} lacks adequate controls",
                f"{dept} processes may have gaps",
                f"{dept} risk assessment requires attention"
            ],
            "risk_types": {
                "Operational": [f"{dept} operational processes need review"],
                "Financial": [f"{dept} financial controls should be evaluated"],
                "Fraud": [f"{dept} fraud prevention needs assessment"],
                "Financial_Fraud": [f"{dept} financial reporting controls need review"],
                "Operational_Fraud": [f"{dept} operational override controls need assessment"]
            },
            "summary": f"The {dept} department shows a {suggested_risk.lower()} overall risk level based on analysis of control objectives and risk categories."
        }

def calculate_risk_score(data: Dict[str, Any]) -> str:
    """Calculate risk score based on risk distribution"""
    try:
        if "risk_distribution" not in data:
            return "N/A"
        
        risk_dist = data["risk_distribution"]
        
        # Calculate weighted score: High=3, Medium=2, Low=1
        high_count = risk_dist.get("High", 0)
        medium_count = risk_dist.get("Medium", 0)
        low_count = risk_dist.get("Low", 0)
        
        total_risks = high_count + medium_count + low_count
        if total_risks == 0:
            return "N/A"
        
        weighted_score = (high_count * 3 + medium_count * 2 + low_count * 1) / total_risks
        
        # Convert to a 0-10 scale
        normalized_score = min(10, max(0, (weighted_score / 3) * 10))
        
        return f"{normalized_score:.1f}"
    
    except Exception as e:
        logger.error(f"Error calculating risk score: {str(e)}")
        return "N/A"

def generate_recommendations(model, data: Dict[str, Any]) -> List[Dict[str, str]]:
    """Generate recommendations based on the analyzed data"""
    try:
        # Prepare data for the prompt
        gaps_str = ""
        for idx, gap in enumerate(data.get("gaps", [])):
            gaps_str += f"{idx+1}. Department: {gap.get('department', '')}\n"
            gaps_str += f"   Control Objective: {gap.get('control_objective', '')}\n"
            gaps_str += f"   Description: {gap.get('description', '')}\n"
            gaps_str += f"   Risk Impact: {gap.get('risk_impact', '')}\n\n"
        
        # If no gaps found, use control objectives
        if not gaps_str and "control_objectives" in data:
            for idx, obj in enumerate(data.get("control_objectives", [])[:5]):  # Limit to 5 for brevity
                gaps_str += f"{idx+1}. Department: {obj.get('department', '')}\n"
                gaps_str += f"   Control Objective: {obj.get('objective', '')}\n"
                gaps_str += f"   What Can Go Wrong: {obj.get('what_can_go_wrong', '')}\n"
                gaps_str += f"   Risk Level: {obj.get('risk_level', '')}\n\n"
        
        if not gaps_str:
            # Return empty recommendations if no data
            return []
        
        prompt = f"""
        As a Risk Control and Audit expert, analyze the following risk control gaps and generate 3-5 strategic recommendations to improve the overall risk management:

        GAPS AND RISKS:
        {gaps_str}

        For each recommendation, provide:
        1. A concise title
        2. Priority level (High, Medium, Low)
        3. A detailed description of the recommendation
        4. Expected impact of implementing the recommendation
        5. Implementation complexity (High, Medium, Low)

        Respond with ONLY a JSON array containing the recommendations:
        [
            {{
                "title": "string",
                "priority": "string",
                "description": "string",
                "impact": "string",
                "complexity": "string"
            }}
        ]
        """
        
        # Get response from Gemini
        response = model.generate_content(prompt)
        
        # Extract JSON from the response
        try:
            response_text = response.text
            # Check if response has JSON code blocks and extract them
            if "```json" in response_text:
                json_text = response_text.split("```json")[1].split("```")[0].strip()
            elif "```" in response_text:
                json_text = response_text.split("```")[1].strip()
            else:
                json_text = response_text.strip()
            
            recommendations = json.loads(json_text)
            return recommendations
        
        except Exception as json_error:
            logger.error(f"Error extracting recommendations JSON from Gemini response: {str(json_error)}")
            logger.debug(f"Raw response: {response.text}")
            return []
    
    except Exception as e:
        logger.error(f"Error generating recommendations: {str(e)}")
        return []

def generate_department_recommendations(model, data: Dict[str, Any]) -> List[Dict[str, str]]:
    """Generate recommendations focused on each department"""
    try:
        # Get list of departments
        departments = data.get("departments", [])
        if not departments:
            return []
            
        # Get department risks
        dept_risks = data.get("department_risks", {})
        
        all_recommendations = []
        
        for dept in departments:
            # Collect department-specific data
            dept_data = dept_risks.get(dept, {})
            
            # Skip if empty
            if not dept_data:
                continue
                
            risk_level = dept_data.get("overall_risk_level", "Medium")
            key_risks = dept_data.get("key_risks", [])
            
            # Create a summary of risk data for this department
            dept_info = f"Department: {dept}\n"
            dept_info += f"Overall Risk Level: {risk_level}\n"
            
            if key_risks:
                dept_info += "Key Risks:\n"
                for risk in key_risks:
                    dept_info += f"- {risk}\n"
                    
            # Get objectives for this department
            dept_objectives = [obj for obj in data.get("control_objectives", []) if obj.get("department") == dept]
            
            if dept_objectives:
                # Find high risk objectives
                high_risk_objs = [obj for obj in dept_objectives if obj.get("risk_level", "").lower() in ["high", "h", "critical"]]
                med_risk_objs = [obj for obj in dept_objectives if obj.get("risk_level", "").lower() in ["medium", "m", "moderate"]]
                
                if high_risk_objs or med_risk_objs:
                    # Prioritize high risk objectives, then medium
                    priority_objs = high_risk_objs[:2] + med_risk_objs[:1] if high_risk_objs else med_risk_objs[:2]
                    
                    for obj in priority_objs:
                        dept_info += f"\nControl Objective: {obj.get('objective', '')}\n"
                        dept_info += f"Risk: {obj.get('what_can_go_wrong', '')}\n"
                        dept_info += f"Risk Level: {obj.get('risk_level', '')}\n"
                        
            # Generate recommendations using Gemini
            prompt = f"""
            As a Risk Control Matrix expert, create detailed, specific recommendations for the {dept} department based on the risk analysis below:
            
            {dept_info}
            
            Create 2-3 specific, actionable recommendations that:
            1. Address the highest-priority risks identified
            2. Provide detailed, practical solutions (not general advice)
            3. Include specific actions, tools, or controls to implement
            4. Explain the expected impact of implementing the recommendation
            
            For each recommendation, include:
            - A clear title summarizing the recommendation
            - A detailed description with specific steps for implementation (at least 3-4 sentences)
            - The expected impact/benefit
            - The priority level (High/Medium/Low)
            
            Format your response as a JSON array:
            [
                {{
                    "department": "{dept}",
                    "title": "Recommendation Title",
                    "description": "Detailed, specific recommendation with actionable steps",
                    "impact": "Expected impact of implementation",
                    "priority": "High/Medium/Low"
                }}
            ]
            """
            
            try:
                # Get response from Gemini
                response = model.generate_content(prompt)
                response_text = response.text
                
                # Extract JSON
                if "```json" in response_text:
                    json_text = response_text.split("```json")[1].split("```")[0].strip()
                elif "```" in response_text:
                    json_text = response_text.split("```")[1].strip()
                else:
                    json_text = response_text.strip()
                
                # Parse recommendations
                dept_recommendations = json.loads(json_text)
                
                # Ensure it's a list
                if isinstance(dept_recommendations, dict):
                    dept_recommendations = [dept_recommendations]
                    
                all_recommendations.extend(dept_recommendations)
                    
            except Exception as e:
                logger.error(f"Error generating recommendations for {dept}: {str(e)}")
                # Add a fallback recommendation
                all_recommendations.append({
                    "department": dept,
                    "title": f"Review Control Framework for {dept}",
                    "description": f"Conduct a comprehensive review of the control framework in the {dept} department, focusing on high-risk areas. Implement additional preventive controls to address potential gaps and automate manual processes where possible to reduce human error.",
                    "impact": "Strengthened control environment and reduced risk exposure",
                    "priority": risk_level
                })
        
        return all_recommendations
                
    except Exception as e:
        logger.error(f"Error generating department recommendations: {str(e)}")
        return [{
            "department": "All Departments",
            "title": "Enhance Risk Control Framework",
            "description": "Conduct a comprehensive review of the control framework across all departments, focusing on high-risk areas. Implement additional preventive controls and automate manual processes where possible.",
            "impact": "Strengthened control environment and reduced risk exposure",
            "priority": "High"
        }] 