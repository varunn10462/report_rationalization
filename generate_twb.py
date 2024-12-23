import random

def generate_sql_queries(num_queries):
    sensitive_queries = [
        "SELECT 'John' AS FIRST_NAME, 'Doe' AS LAST_NAME, '1234567890123456' AS ACCOUNT_NUMBER, '123 Elm St, Apt 4B, Springfield, IL, 62704' AS ADDRESS, 'Premium Member' AS MEMBER_STATUS, '1985-08-15' AS DOB, 'Male' AS GENDER, '555-123-4567' AS PHONE_NUMBER, 'john.doe@email.com' AS EMAIL, '4111111111111111' AS CARD_NUMBER, '123-45-6789' AS SSN, '12-3456789' AS TIN FROM DUAL",
        "SELECT 'Jane' AS FIRST_NAME, 'Smith' AS LAST_NAME, '9876543210987654' AS ACCOUNT_NUMBER, '456 Oak St, Apt 7C, Springfield, IL, 62704' AS ADDRESS, 'Regular Member' AS MEMBER_STATUS, '1990-12-25' AS DOB, 'Female' AS GENDER, '555-987-6543' AS PHONE_NUMBER, 'jane.smith@email.com' AS EMAIL, '4222222222222222' AS CARD_NUMBER, '987-65-4321' AS SSN, '98-7654321' AS TIN FROM DUAL"
    ]
    
    non_sensitive_queries = [
        "SELECT id, name, created_at FROM products",
        "SELECT id, title, content FROM posts"
    ]
    
    queries = []
    for _ in range(num_queries):
        if random.choice([True, False]):
            queries.append(random.choice(sensitive_queries))
        else:
            queries.append(random.choice(non_sensitive_queries))
    
    return queries

sql_queries = generate_sql_queries(300)

# Generate the XML content for the .twb file
xml_content = """<?xml version='1.0' encoding='utf-8' ?>
<workbook>
  <datasources>
    <datasource>
      <connection>
        <relation type="text">
          SELECT * FROM users WHERE email = 'example@example.com'
        </relation>
      </connection>
      <connection>
        <relation type="text">
          SELECT * FROM orders WHERE order_date = '2023-01-01'
        </relation>
      </connection>
      <connection>
        <relation type="text">
          SELECT 
            'John' AS FIRST_NAME, 
            'Doe' AS LAST_NAME, 
            '1234567890123456' AS ACCOUNT_NUMBER, 
            '123 Elm St, Apt 4B, Springfield, IL, 62704' AS ADDRESS, 
            'Premium Member' AS MEMBER_STATUS, 
            '1985-08-15' AS DOB, 
            'Male' AS GENDER, 
            '555-123-4567' AS PHONE_NUMBER, 
            'john.doe@email.com' AS EMAIL, 
            '4111111111111111' AS CARD_NUMBER, 
            '123-45-6789' AS SSN, 
            '12-3456789' AS TIN
          FROM DUAL
        </relation>
      </connection>
"""

for query in sql_queries:
    xml_content += f"""
      <connection>
        <relation type="text">
          {query}
        </relation>
      </connection>
"""

xml_content += """
    </datasource>
  </datasources>
</workbook>
"""

# Save the XML content to the .twb file
with open("C:\\Users\\varunsharma\\Desktop\\Sanu_task_version_2\\final doc\\ADS_Channel_Summary_Weekly.twb", "w") as file:
    file.write(xml_content)