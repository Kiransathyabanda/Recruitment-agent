# Recruitment-agent
Recruitment agent using CrewAI and the data will be saved into an excel file.

# Resume Processing Pipeline

This repository contains a resume processing pipeline that extracts information from PDF resumes, categorizes them as 'Fresher' or 'Experienced', extracts relevant attributes, assigns scores for shortlisting, and saves the extracted data into an Excel file.

## Features

- Extracts text from PDF resumes.
- Categorizes resumes as 'Fresher' or 'Experienced'.
- Extracts key attributes such as name, contact details, skills, projects, etc.
- Assigns scores to resumes based on predefined criteria for shortlisting.
- Saves the extracted data into an Excel file.

## Installation

To set up the project, you need to have Python installed. Follow the steps below to create a virtual environment and install the required packages.

1. Clone the repository:

    ```sh
    git clone https://github.com/your-username/resume-processing-pipeline.git
    cd resume-processing-pipeline
    ```

2. Create a virtual environment:

    ```sh
    python -m venv venv
    ```

3. Activate the virtual environment:

    - On Windows:

        ```sh
        venv\Scripts\activate
        ```

    - On macOS/Linux:

        ```sh
        source venv/bin/activate
        ```

4. Install the required packages:

    ```sh
    pip install -r requirements.txt
    ```

## Usage

To process the resumes and extract the information, follow these steps:

1. Place your PDF resumes in the `resumes` directory.

2. Run the main script:

    ```sh
    python main.py
    ```

3. The results will be saved in the `resume_data.xlsx` file in the root directory.

## File Structure
resume-processing-pipeline/
- │
- ├── resumes/ # Directory containing PDF resumes
- ├── main.py # Main script to process the resumes
- ├── requirements.txt # List of required packages
- ├── README.md # Project documentation
- ├── tools.py # Tool functions used in the pipeline
- └── write_to_excel_tool.py # Function to write data to Excel file

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any changes or improvements.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
