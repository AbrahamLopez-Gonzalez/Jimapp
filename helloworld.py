from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

class ContactManager:
    def __init__(self, filename="contacts.xlsx"):
        self.filename = filename
        self.workbook = load_workbook(filename)
        self.sheet = self.workbook.active
        self.add_headers()

    def add_headers(self):
        headers = ['Name', 'Phone', 'Email']
        for col, header in enumerate(headers, start=1):
            if self.sheet.cell(row=1, column=col).value is None:
                self.sheet.cell(row=1, column=col, value=header).font = Font(bold=True)

    def add_contact(self, name, phone, email):
        next_row = self.sheet.max_row + 1
        self.sheet.cell(row=next_row, column=1, value=name)
        self.sheet.cell(row=next_row, column=2, value=phone)
        self.sheet.cell(row=next_row, column=3, value=email)
        self.workbook.save(self.filename)

    def list_contacts(self):
        for row in range(2, self.sheet.max_row + 1):
            name = self.sheet.cell(row=row, column=1).value
            phone = self.sheet.cell(row=row, column=2).value
            email = self.sheet.cell(row=row, column=3).value
            print(f"Name: {name}, Phone: {phone}, Email: {email}")

# Example usage
contact_manager = ContactManager()

# Adding contacts
contact_manager.add_contact("John Doe", "1234567890", "john@example.com")
contact_manager.add_contact("Jane Smith", "0987654321", "jane@example.com")

# Listing contacts
print("Listing all contacts:")
contact_manager.list_contacts()