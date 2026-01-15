# AI-Based-Risk-Analysis-for-Industrial-Maintenance-Procedures
Automated risk analysis for industrial maintenance procedures using Python, AI, and Excel integration.

This project automates industrial maintenance risk analysis by leveraging artificial intelligence. It processes maintenance procedures from Excel files, utilizing the OpenAI API with a structured technical prompt to assess potential hazards, classify risk probability and severity, and recommend control measures.
Each procedure is individually analyzed through a looped pipeline, with outputs standardized and validated via helper functions. A risk matrix is then applied to determine the final risk classification: Trivial, Tolerable, Substantial, or Intolerable.
Results are written back to Excel with conditional color formatting, providing an intuitive visual overview of risk levels for rapid review and decision-making.


Key Features:
AI-Powered Analysis: Uses structured prompts to ensure consistent, contextual risk assessment
Excel Integration: Reads from and writes to existing Excel templates, preserving workflow compatibility
Automated Risk Matrix: Applies industry-standard severity/probability classification
Visual Output: Color-coded formatting for immediate risk recognition
Scalable Processing: Handles multiple procedures through automated batch processing


Technology Stack: Python, OpenAI GPT-4 API, Pandas, Openpyxl, Industrial Risk Management Standards
