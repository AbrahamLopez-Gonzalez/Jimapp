from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

class ContactManager:
    def __init__(self, filename):
        self.filename = filename
        try:
            self.workbook = load_workbook(filename)
        except FileNotFoundError:
            self.workbook = Workbook()
        self.sheet = self.workbook.active

        # Define headers if they don't exist
        headers = ["Name", "Phone", "Email"]
        if self.sheet.max_row == 0 or any(self.sheet.cell(row=1, column=col+1).value != header
                                          for col, header in enumerate(headers)):
            self.add_headers(headers)

    def add_headers(self, headers):
        for col, header in enumerate(headers, start=1):
            cell = self.sheet.cell(row=1, column=col)
            cell.value = header
            cell.font = Font(bold=True)

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
contact_manager = ContactManager("contacts.xlsx")

contact_manager.add_contact("John Doe", "123-456-7890", "john@example.com")
contact_manager.add_contact("Jane Smith", "987-654-3210", "jane@example.com")

print("Contacts:")
contact_manager.list_contacts()