from flask import Flask, render_template, request
import time

app = Flask(__name__)

# Define a function to ask questions to the customer
def ask_questions(name, product_choice, issue_category_choice, issue_subcategory_choice, machine_area):
    # Define a list of available products
    products = ["Machine type X", "Machine type Y", "Service type A", "Service type B"]
    selected_product = products[product_choice-1]
  
    # Define a list of available issue categories
    issue_categories = ["Technical issue", "Billing issue", "Account issue"]
    selected_issue_category = issue_categories[issue_category_choice-1]
  
    # Define a list of available issue sub-categories
    issue_subcategories = ["Electrical issue", "Mechanical issue", "Software issue", "Other issue"]
    selected_issue_subcategory = issue_subcategories[issue_subcategory_choice-1]
  
    # Determine the root cause of the issue based on the customer's answers
    root_cause = f"The issue is related to {selected_product}, and it is a {selected_issue_category.lower()} issue, specifically a {selected_issue_subcategory.lower()} in the {machine_area} area."
  
    # Assign the issue to the relevant person based on the root cause
    if selected_issue_subcategory.lower() == "electrical issue":
        person_to_fix = "John Smith (Electrical Engineer)"
        contact_number = "555-1234"
    elif selected_issue_subcategory.lower() == "mechanical issue":
        person_to_fix = "Mary Johnson (Mechanical Engineer)"
        contact_number = "555-5678"
    elif selected_issue_subcategory.lower() == "software issue":
        person_to_fix = "David Lee (Software Engineer)"
        contact_number = "555-9876"
    else:
        person_to_fix = "Jane Doe (Technical Support)"
        contact_number = "555-4321"
  
    # Send a message to the relevant person to fix the issue
    return f"Thank you for your answers, {name}! Based on your responses, the root cause of the issue is: {root_cause}. I will notify {person_to_fix} to fix the issue. You can contact them at {contact_number}. Thank you for using our chatbot!"


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Get the customer's answers from the form
        name = request.form['name']
        product_choice = int(request.form['product'])
        issue_category_choice = int(request.form['issue_category'])
        issue_subcategory_choice = int(request.form['issue_subcategory'])
        machine_area = request.form['machine_area']
    
        # Call the ask_questions function to determine the root cause and assign the issue to the relevant person
        message = ask_questions(name, product_choice, issue_category_choice, issue_subcategory_choice, machine_area)
      
        return render_template('result.html', message=message)
    else:
        # Render the chatbot form
        products = ["Machine type X", "Machine type Y", "Service type A", "Service type B"]
        issue_categories = ["Technical issue", "Billing issue", "Account issue"]
        issue_subcategories = ["Electrical issue", "Mechanical issue", "Software issue", "Other issue"]
        return render_template('index.html', products=products, issue_categories=issue_categories, issue_subcategories=issue_subcategories)


if __name__ == '__main__':
    app.run(debug=True)
