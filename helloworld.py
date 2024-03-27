from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
#We need openpyxl library to be able to create, read and write on excel
#.styles is used in this case to manipulate if the Font is bold or normal

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

    def edit_contact(self, name, new_phone, new_email):
        for row in range(2, self.sheet.max_row + 1):
            if self.sheet.cell(row=row, column=1).value == name:
                self.sheet.cell(row=row, column=2, value=new_phone)
                self.sheet.cell(row=row, column=3, value=new_email)
                self.workbook.save(self.filename)
                print(f"Contact '{name}' updated successfully.")
                return
        print(f"Contact '{name}' not found.")

    def search_contact(self, name):
        for row in range(2, self.sheet.max_row + 1):
            if self.sheet.cell(row=row, column=1).value == name:
                phone = self.sheet.cell(row=row, column=2).value
                email = self.sheet.cell(row=row, column=3).value
                print(f"Name: {name}, Phone: {phone}, Email: {email}")
                return
        print(f"Contact '{name}' not found.")

    def delete_contact(self, name):
        for row in range(2, self.sheet.max_row + 1):
            if self.sheet.cell(row=row, column=1).value == name:
                self.sheet.delete_rows(row, amount=1)
                self.workbook.save(self.filename)
                print(f"Contact '{name}' deleted successfully.")
                return
        print(f"Contact '{name}' not found.")

# Example usage
contact_manager = ContactManager("contacts.xlsx")

while True:
    print("\nContact Manager")
    print("1. Add Contact")
    print("2. Edit Contact")
    print("3. Search Contact")
    print("4. Delete Contact")
    print("5. List Contacts")
    print("6. Exit")
    choice = input("Enter your choice (1-6): ")

    if choice == "1":
        name = input("Enter name: ")
        phone = input("Enter phone: ")
        email = input("Enter email: ")
        contact_manager.add_contact(name, phone, email)
    elif choice == "2":
        name = input("Enter name of contact to edit: ")
        new_phone = input("Enter new phone: ")
        new_email = input("Enter new email: ")
        contact_manager.edit_contact(name, new_phone, new_email)
    elif choice == "3":
        name = input("Enter name to search: ")
        contact_manager.search_contact(name)
    elif choice == "4":
        name = input("Enter name to delete: ")
        contact_manager.delete_contact(name)
    elif choice == "5":
        contact_manager.list_contacts()
    elif choice == "6":
        break
    else:
        print("Invalid choice. Please try again.")