
# Import the required libraries
from tkinter import *
from tkinter import ttk, filedialog
import pandas as pd
from tkinter.messagebox import showinfo
import sys
# Create an instance of tkinter frame
win = Tk()

# Set the size of the tkinter window
win.title("Data Convert for Candy, Viri")
win.geometry("450x100")


def close():
    sys.exit()


def open_file():
    filename = filedialog.askopenfilename(title="Open a File", filetype=(
        ("xlxs files", ".*xlsx"), ("All Files", "*.")))
    # show the open file dialog
    filename = r"{}".format(filename)
    sheet = pd.read_excel(filename)

    if sheet is not None:
        out_df = pd.DataFrame(
            columns=['answer_text', 'question', 'is_correct'])

        questions = []
        answer_texts = []
        is_corrects = []
        for index, row in sheet.iterrows():
            questions.append(row['question'])
            answer_texts.append(row['T'])
            is_corrects.append(1)
            questions.append(row['question'])
            answer_texts.append(row['F'])
            is_corrects.append(0)
            questions.append(row['question'])
            answer_texts.append(row['F.1'])
            is_corrects.append(0)

            if pd.notna(row['F.2']):
                questions.append(row['question'])
                answer_texts.append(row['F.2'])
                is_corrects.append(0)

        out_df['answer_text'] = answer_texts
        out_df['question'] = questions
        out_df['is_correct'] = is_corrects

        out_df.to_excel(".\\result\\Answer Options.xlsx")
        # out_df.to_excel('.result\Answer Options.xlsx')
        print(out_df)

        showinfo(
            title='Data Convert',
            message='Hi, Candy, Answer Options.xlsx is generated or updated.'
        )


# Open file button
open_button = ttk.Button(
    win,
    text='Open a File',
    command=open_file
)
# Open file button
close_button = ttk.Button(
    win,
    text='Close',
    command=close
)

# Display button for the window
open_button.grid(column=1, row=1, sticky='w', padx=10, pady=10)
close_button.grid(column=3, row=1, padx=10, pady=10)

# open_button.pack(expand=True)
# close_button.pack(expand=True)

win.mainloop()
