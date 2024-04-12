from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
# We need openpyxl library to be able to create, read and write on excel
# .styles is used in this case to manipulate if the Font is bold or normal

class ContactManager:
    def __init__(self, filename):
        self.filename = filename
        try:
            self.workbook = load_workbook(filename)
        except FileNotFoundError:
            self.workbook = Workbook()
        self.sheet = self.workbook.active

        # Define headers if they don't exist
        headers = ["F.Name", "L.Name", "Phone", "Birthday"]
        if self.sheet.max_row == 0 or any(self.sheet.cell(row=1, column=col+1).value != header
                                          for col, header in enumerate(headers)):
            self.add_headers(headers)

    #If headers already exist
    def add_headers(self, headers):
        for col, header in enumerate(headers, start=1):
            cell = self.sheet.cell(row=1, column=col)
            cell.value = header
            cell.font = Font(bold=True)

    def add_contact(self, first_name, last_name, phone, birthday):
        next_row = self.sheet.max_row + 1
        self.sheet.cell(row=next_row, column=1, value=first_name)
        self.sheet.cell(row=next_row, column=2, value=last_name)
        self.sheet.cell(row=next_row, column=3, value=phone)
        self.sheet.cell(row=next_row, column=4, value=birthday)
        self.workbook.save(self.filename)

    def list_contacts(self):
        for row in range(2, self.sheet.max_row + 1):
            first_name = self.sheet.cell(row=row, column=1).value
            last_name = self.sheet.cell(row=row, column=2).value
            phone = self.sheet.cell(row=row, column=3).value
            birthday = self.sheet.cell(row=row, column=4).value
            print(f"First Name: {first_name}, Last Name: {last_name}, Phone: {phone}, Birthday: {birthday}")

    def edit_contact(self, first_name, new_phone, new_birthday):
        for row in range(2, self.sheet.max_row + 1):
            if self.sheet.cell(row=row, column=1).value == first_name:
                self.sheet.cell(row=row, column=3, value=new_phone)
                self.sheet.cell(row=row, column=4, value=new_birthday)
                self.workbook.save(self.filename)
                print(f"Contact '{first_name}' updated successfully.")
                return
        print(f"Contact '{first_name}' not found.")

    def search_contact(self, first_name):
        for row in range(2, self.sheet.max_row + 1):
            if self.sheet.cell(row=row, column=1).value == first_name:
                last_name = self.sheet.cell(row=row, column=2).value
                phone = self.sheet.cell(row=row, column=3).value
                birthday = self.sheet.cell(row=row, column=4).value
                print(f"Name: {first_name}, Last Name: {last_name}, Phone: {phone}, Birthday: {birthday}")
                return
        print(f"Contact '{first_name}' not found.")

    def delete_contact(self, first_name):
        for row in range(2, self.sheet.max_row + 1):
            if self.sheet.cell(row=row, column=1).value == first_name:
                self.sheet.delete_rows(row, amount=1)
                self.workbook.save(self.filename)
                print(f"Contact '{first_name}' deleted successfully.")
                return
        print(f"Contact '{first_name}' not found.")

    #Adding new functions to validate for user input 
    # def phone_validation(self,phone):
    #     if phone != int:#Validate input with the value range iterating through the string
    #         print(f"The number: {phone} is not a valid number.\nRember, a phone number contains 10 digits.")
    #         return

# usage
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

    match choice:
        case "1":
    # if choice == "1":
            first_name = input("Enter first name: ")
            # don't pass until != empty
            last_name = input("Enter last name: ")
            # don't pass until != empty
            phone = input("Enter phone: ")
            # don't pass until != empty
            birthday = input("Enter birthday: ")
            # don't pass until != empty
            # contact_manager.phone_validation(phone)
            contact_manager.add_contact(first_name, last_name, phone, birthday)
    # elif choice == "2":
        case "2":
            first_name = input("Enter first name of contact to edit: ")
            # only accept fist name
            new_phone = input("Enter new phone: ")
            new_birthday = input("Enter new birthday: ")
            contact_manager.edit_contact(first_name, new_phone, new_birthday)
    # elif choice == "3":
        case "3":
            first_name = input("Enter first name to search: ")
            # enter the lastname or add the wildcard search
            contact_manager.search_contact(first_name)
    # elif choice == "4":
        case "4":
            first_name = input("Enter first name to delete: ")
            contact_manager.delete_contact(first_name)
    # elif choice == "5":
        case "5":
            contact_manager.list_contacts()
    # elif choice == "6":
        case "6":
            break
    # else:
        case _:
            print("Invalid choice. Please try again.")