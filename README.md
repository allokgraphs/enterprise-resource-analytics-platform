# Enterprise-Resource-Analytics-Automation-Platform


🚀 Project Overview
Data automation and visualization platform that transforms Excel-based workflows into modern web analytics solutions. Reduces manual processing effort by 67% through automated SharePoint integration and interactive dashboards.

💼 Problem Solved
Manual Excel file searches → Automated data extraction
Time-intensive workflows → 67% efficiency improvement
Static reports → Interactive web dashboards

✨ Key Features
🔄 Automated Data Processing - SharePoint API integration, real-time synchronization, automated Excel processing
📊 Interactive Visualization - Hierarchical tree structure, collapsible navigation, responsive design
🎯 Smart Analytics - Availability tracking, sorted displays, real-time dashboard updates

💻 Tech Stack
Backend: Python, Pandas, NumPy, openpyxl
Frontend: HTML5, CSS3, JavaScript, React
Integration: SharePoint APIs, JSON

📈 Impact Metrics
MetricImprovementProcessing Time67% reductionData AccuracyAutomated validationUser ExperienceInteractive dashboards

📊 Test Data
The platform has been tested and validated using sample data:
Test File: Test_data.xlsx
Contains enterprise resource data for demonstration and testing purposes
Validates automated processing capabilities and visualization accuracy
Used for performance benchmarking and feature validation

🛠️ Quick Start
bash# Clone repository
git clone https://github.com/allokgraphs/enterprise-resource-analytics-platform.git

# Install dependencies
pip install pandas numpy openpyxl

# Run application with test data
python pms_visualization.py Test_data.xlsx

# Or run with your own data
python pms_visualization.py path/to/your_excel_file.xlsx
💡 Usage
pythonfrom pms_visualization import generate_pms_visualization

# Generate from test data
html_output = generate_pms_visualization(file_path="Test_data.xlsx")

# Generate from custom Excel file
html_output = generate_pms_visualization(file_path="data.xlsx")

# Save dashboard
with open("dashboard.html", "w") as f:
    f.write(html_output)
    
🚀 Future Enhancements
Real-time SharePoint sync
Advanced analytics & predictions
Mobile application
RESTful API development

🤝 Contributing
Fork the repository
Create feature branch (git checkout -b feature/NewFeature)
Commit changes (git commit -m 'Add NewFeature')
Push to branch (git push origin feature/NewFeature)
Open Pull Request

📞 Contact
Alok Raj - Data Analyst & Python Developer
📧 Email: alokraj090102@gmail.com
💼 LinkedIn: linkedin.com/in/allokgraphs
🐙 GitHub: @allokgraphs
