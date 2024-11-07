import os
def get_names():
    try:
        with open("mail merge/Input/Names/inivited_names.txt","r") as file:
            names_list = file.readlines()
            return names_list
    except Exception:
        print("Failed to open names file")


def main():
    try:
        with open("mail merge/Input/letters/starting_letter.docx", 'r') as file:
            text = file.readlines()
    except Exception:
        print("Failed to open starting letter")

    directory_name = "ReadyToSend"

    try:
        os.mkdir(f"mail merge/Input/letters/{directory_name}")
    except Exception:
        print("Failed to create directory")

    names = get_names()
    for name in names:
        name = name.replace("\n","")
        try:
            with open(f"mail merge/Input/letters/ReadyToSend/{name}.docx","w") as letter:
                for line in text:
                    line = line.replace("[name]", name)
                    if len(line) > 2:
                        line = line.replace("\n","")
                    letter.write(line)
        except Exception:
            print("Failed to write new letter")
if __name__ == "__main__":
    main()