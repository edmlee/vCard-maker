from openpyxl import load_workbook


class Contact():
    def __init__(self, first_name, last_name, phone, id):
        self.first_name = first_name
        self.last_name = last_name
        self.phone = phone
        self.id = id


def clear_vcard(device):
    try:
        file = open(f"./Contacts/Contacts {device}.vcf", "w")
        file.close()
    except FileNotFoundError:
        print(f"Missing \"{folder_name}\" folder in current directory")
        quit()


def print_vcard(contact):
    file.write("BEGIN:VCARD\n")
    file.write("VERSION:2.1\n")
    file.write(f"N:{contact.last_name};{contact.first_name};;;\n")
    file.write(f"FN:{contact.first_name} {contact.last_name}\n")
    file.write(f"TEL;CELL:{contact.phone}\n")
    file.write("END:VCARD\n")


# Main
wb = load_workbook("contact_list.xlsx")
ws = wb["Contacts"]
folder_name = "Contacts" # Folder must be created in the main directory. Used to store vCards
device_col = 1
first_name_col = 2
last_name_col = 3
phone_col = 4

num_of_contacts = 14 # Number of contacts to add per vCard
num_of_devices = int((ws.max_row - 1) // num_of_contacts)
row = 2

for i in range(num_of_devices):
    device = int(ws.cell(row, device_col).value)
    clear_vcard(device)

    for j in range(row, row + num_of_contacts):
        try:
            with open(f"./{folder_name}/Contacts {device}.vcf", "a") as file:
                first_name = ws.cell(j, first_name_col).value
                last_name = ws.cell(j, last_name_col).value
                phone = int(ws.cell(j, phone_col).value)
                contact = Contact(first_name, last_name, phone, device)
                print_vcard(contact)
        except FileNotFoundError:
            print(f"Missing \"{folder_name}\" folder in current directory")
            quit()
    row += num_of_contacts
