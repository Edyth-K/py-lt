from datetime import datetime
import shutil
import os

def main():
    today = datetime.now()
    file_name = (today.strftime("%m-%d-%Y"))

    source_file = "template.md"
    new_title = f"📅 Date: {file_name}"
    new_filename = f"{file_name}.md"

    # Step 0: Check if file already exists
    if os.path.exists(new_filename):
        print(f"File '{new_filename}' already exists. Aborting.")
    else:
        # Step 1: Copy the file
        shutil.copyfile(source_file, new_filename)

        # Step 2: Read the file, replace the title
        with open(new_filename, "r+") as f:
            lines = f.readlines()
            if lines:
                lines[0] = f"# {new_title}\n"  # Replace first line with new title
            else:
                lines = [f"# {new_title}\n"]
            f.seek(0)
            f.writelines(lines)
            f.truncate()

        print(f"Created '{new_filename}' successfully.")

if __name__ == "__main__":
    main()