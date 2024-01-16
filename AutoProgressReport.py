import pandas as pd

def generate_responses(student_data):
    # Process student data and generate responses
    responses = []

    for student in student_data:
        name, strengths, problem_areas, address_strategy, concerns = student.split('/')
        response = f"""
        1. **Student Name:** {name}

        2. **Question 1 - Strengths:**
           - {strengths}

        3. **Question 2 - Problem Areas:**
           - {problem_areas}

        4. **Question 3 - How to Address Problem Areas Next Month:**
           - {address_strategy}

        5. **Question 4 - Concerns:**
           - {concerns}
        """
        responses.append(response)

    return responses

def main():
    # Read student data from Excel file
    df = pd.read_excel('your_spreadsheet.xlsx', header=None)

    # Generate responses
    responses = generate_responses(df[0])

    # Print or save responses as needed
    for response in responses:
        print(response)

if __name__ == "__main__":
    main()
