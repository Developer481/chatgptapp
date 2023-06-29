import gspread
import openai

# Set up your OpenAI API credentials
openai.api_key = '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'

# Function to generate keywords
def generate_keywords(num_keywords):
    # Generate keywords using the Python ChatGPT API
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt= num_keywords,
        max_tokens=100,
        n=1,
        stop=None,
        temperature=0.3,
        top_p=1.0,
        frequency_penalty=0.0,
        presence_penalty=0.0
    )

    # Extract and clean up the generated keywords
    keywords = response.choices[0].text.strip().split(',')
    keywords = [keyword.strip() for keyword in keywords]

    # Remove empty and duplicate keywords
    keywords = list(filter(None, keywords))
    keywords = list(set(keywords))

    return keywords

# Number of keywords to generate
#num_keywords = 10

sa = gspread.service_account(filename="C://Users/EDITOR/Desktop/Googleapp/developer-374205-27d86b37709f.json")
sheet = sa.open("sample_sales")
work_sheet = sheet.worksheet("products_2023-06-23_12_20_35")

#find sheet last row 


# Get the values in the 3rd column
column_3_values = work_sheet.col_values(3)

# Find the last non-empty cell in the 3rd column
start = len(column_3_values)
start1 = start+1


# Get the values in the 2nd column
column_2_values =work_sheet.col_values(2)

# Find the last non-empty cell in the 2nd column
end = len(column_2_values)





for i in range(start1 ,end+1):
    try:
        v1 = work_sheet.cell(i,2).value
        v2  = "write 10 comma-separated keywords relating to "
        v4 = "write 30 comma-separated keywords relating to "
        v3 = v2+""+v1
        v5 = v2+""+v4
        # Generate and print the keywords
        keywords10 = generate_keywords(v3)
        keywords30 = generate_keywords(v5)
        
        data = ','.join(keywords10)
        data1 = ','.join(keywords30)
        
        result = ''.join(char for char in data if not char.isdigit())
        result1 = ''.join(char for char in data1 if not char.isdigit())
        
        ds = "C"
        ds1 = "D"
        
        cell_range = ds+str(i)
        cell_range1 = ds1+str(i)
        
        # Update the values in the worksheet
        work_sheet.update(cell_range, result)
        work_sheet.update(cell_range1,result1)
        
        
        
        
    except:
        # Update the values in the worksheet
        work_sheet.update(cell_range, 'Invalid title')
        work_sheet.update(cell_range1,'Invalid title')
        
        continue
        
