from pms_visualization import generate_pms_visualization
 
def main():
    try:
        # Replace with your actual file path
        excel_file = "Test_data.xlsx" #this is the test file 
        
        # Generate the visualization
        html_content = generate_pms_visualization(file_path=excel_file)
        
        # Save to HTML file
        output_file = "pms_dashboard.html"
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(html_content)
        
        print(f"✅ Dashboard successfully generated: {output_file}")
        print("Open the HTML file in your web browser to view the dashboard")
        
    except Exception as e:
        print(f"❌ Error: {e}")
 
if __name__ == "__main__":
    main()
