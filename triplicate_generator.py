import subprocess
import importlib.util
import json
import sys
import os
import csv
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import re
from num2words import num2words

def install(package_name, module_name=None):
    try:
        if module_name:
            package_spec = importlib.util.find_spec(package_name)
            if package_spec and module_name in dir(importlib.import_module(package_name)):
                pass
            else:
                raise ImportError
        else:
            importlib.import_module(package_name)
    except (ImportError, ModuleNotFoundError):
        subprocess.check_call(["pip", "install", package_name])

def import_modules():
    modules_needed = ["num2words", "weasyprint"]
    for module in modules_needed:
        install(module)

def get_data_file_path(file_name): #creating a presistence file in home directory
    user_home = os.path.expanduser("~")
    app_data_dir = os.path.join(user_home, ".mPersistence")
    os.makedirs(app_data_dir, exist_ok=True)
    data_file_path = os.path.join(app_data_dir, file_name)
    return data_file_path

def save_data(data):
    file_path = get_data_file_path("datafrompy.json")
    with open(file_path, "w") as json_file:
        json.dump(data, json_file, indent=4)

def load_data():
    data_file_path = get_data_file_path("datafrompy.json")
    try:
        with open(data_file_path, "r") as json_file:
            return json.load(json_file)
    except FileNotFoundError:
        num_of_rows.insert(0, default_rows)
        return {}  # Return an empty dictionary if the file is not found

def wordnum(number):
    integer_part = int(number)
    decimal_part = number - integer_part
    word=""
    word += num2words(integer_part, lang="en")
    if decimal_part > 0:
        word += " and half" #modify this if you need precise decimal in words
    else:
        pass
    return word

def process_rows(text, pass_marks, number_of_rows):
    unsorted_lines = text.split("\n")
    def sort_key(line):
        parts = re.split(r'[\t ]+', line)
        return int(parts[0])
    numbers_set = set()
    duplicates_found = False

    for line in unsorted_lines:
        parts = re.split(r'[\t ]+', line)
        first_number = int(parts[0])
        second_part = parts[1] if len(parts) > 1 else ""
        full = pass_marks * 2
        if second_part.isdigit() and float(second_part) > float(full):
            show_toast(f"Code {first_number} got {second_part}, which exceeds full mark!")
            # error_label.config(text=f"Code {first_number} got {second_part}, which exceeds full mark!")
            return None

        if first_number in numbers_set:
            duplicates_found = True
            show_toast(f"Duplicate code number: {str(first_number)}")
            # error_label.config(text="Duplicate code number: "+str(first_number))
        else:
            numbers_set.add(first_number)

    if not duplicates_found: # Sort the lines only if no duplicates are found
        lines = sorted(unsorted_lines, key=sort_key)
        line_count = len(lines)
        if (line_count % number_of_rows) != 0:
            pages = (line_count//number_of_rows) + 1
        else:
            pages = (line_count//number_of_rows)
        wordnumber = ""
        row1 = ""
        row2 = ""
        i=0
        for line in lines:
            parts = line.split()  # Split the line using whitespace as the separator
            if len(parts) == 2:  # Ensure that there are exactly two parts (code and marks)
                code, marks = parts[0].strip(), parts[1].strip()
                highlight_css = ""
                try:
                    mark = float(marks)
                    if pass_marks > mark:
                        highlight_css = "highlight"
                    else:
                        highlight_css = ""
                    wordnumber = wordnum(mark)
                except:
                    wordnumber = marks
                    highlight_css = "highlight"
                row1 +=f'''
                    <tr>
                        <td class="bold huge codecolumn">{code}</td>
                        <td class="mocolumn {highlight_css}">{marks}</td>
                    </tr>
                '''
                row2 += f'''
                    <tr>
                        <td class="bold huge codecolumn">{code}</td>
                        <td class="mocolumn {highlight_css}">{marks}</td>
                        <td class="wordnumcolumn {highlight_css}">{wordnumber.capitalize()}</td>
                    </tr>
                '''
        return row1, row2, pages

def generate_pdf(combined_html_content, pdf_file, switch):
    global pdf_created
    from weasyprint import HTML
    if switch == 1:
        with open(pdf_file, 'wb') as f:
            HTML(string=combined_html_content).write_pdf(f)
        pdf_created = pdf_created + 1
        if pdf_created != 0 :
            show_toast(f"Success creating PDF file named: {str(pdf_file)}")
        else:
            show_toast(f"PDF creation failed")
    else: # For development/ debugging only
        try:
            with open("python_output.html", "w") as file:
                file.write(combined_html_content)
            with open(pdf_file, 'xb') as f:
                HTML(string=combined_html_content).write_pdf(f)
            print(f"PDF saved as: {pdf_file}")
            pdf_created += 1
            show_toast(f"Success creating PDF file named: {str(pdf_file)}")
        except FileExistsError:
            os.remove(pdf_file)
            generate_pdf(combined_html_content, pdf_file, switch)

def splice_pages(row_list, pass_marks, number_of_rows):
    content_items = row_list[:number_of_rows] # Extract the first n items from the list
    total_in_page = len(content_items)    #count pass and fail if pass_marks is not None, return them at last
    pass_counts = 0
    if pass_marks is None:
        pass_counts = None
        fail_counts = None
    else:
        for entry in content_items:
            numbers = re.findall(r'\d+', entry)
            number = numbers[1] if len(numbers) > 1 else None
            if number is None:
                pass
            else:
                try:
                    num = float(number)
                    if num >= float(pass_marks):
                        pass_counts += 1
                except ValueError:
                    pass
        fail_counts = total_in_page - pass_counts
    content = "\n".join(["<tr class='rows'>" + item + "</tr>" for item in content_items])
    chopped_row_list = row_list[number_of_rows:]     # Remove the extracted items from the original list
    return content, chopped_row_list, pass_counts, fail_counts

def show_toast(message):
    toast = tk.Toplevel()
    toast.overrideredirect(True)  # Remove window decorations
    label = tk.Label(toast, text=message, bg='lightyellow', fg='red', padx=10, pady=5)
    label.pack()
    toast.lift()
    toast.after(2000, toast.destroy)

def submit_data(switch): # get data, save it to file, create html
    data = {
        "date": date_entry.get(),
        "level": level_entry.get(),
        "program": program_entry.get(),
        "year": year_entry.get(),
        "subject": subject_entry.get(),
        "paper": paper_entry.get(),
        "full_marks": full_marks_entry.get(),
        "type_of_exam": exam_type_var.get(),
        "name": name_entry.get(),
        "num_of_rows": num_of_rows.get(),
        "campus": campus.get(),
        "text_data": text_entry.get("1.0", tk.END).strip(),
    }
    save_data(data)
    if exam_type_var.get() == "Theory":
        thpr = "Th."
    elif exam_type_var.get() == "Practical":
        thpr = "Pr."
    try:
        pass_marks = "{:g}".format(int(data["full_marks"]) / 2)
    except (KeyError, ValueError):
        show_toast("Enter a valid number in full marks!")
    number_of_rows=int(num_of_rows.get())
    global num2words
    combined_html_content = ""
    num = float(pass_marks)
    row1, row2, pages = process_rows(data["text_data"], num, number_of_rows)
    row1_list_unfiltered = re.split(r'<tr>|</tr>', row1)
    row1_list = [part.strip() for part in row1_list_unfiltered if part.strip()]
    row2_list_unfiltered = re.split(r'<tr>|</tr>', row2)
    row2_list = [part.strip() for part in row2_list_unfiltered if part.strip()]
    combined_html_content += f'''
        <!DOCTYPE html>
        <html>
        <head>
        <style>
            @page {{
                size: A4;
                margin: 0.2cm 0.5cm 0.2cm 0.5cm;
            }}
            body {{
                font-family: Arial, sans-serif;
                font-size: 14px;
            }}
            .times {{
                font-family: "Times New Roman", Times, serif;
            }}
            .wholepage {{
                width:100%;
            }}
            .wholepage .mini {{
                font-size:12px;
            }}
            .wholepage tr td table tr * {{
                padding: 2px 0px -2px 0px;
                border: 1px solid;
            }}
            .wholepage .mark-tables td {{
                text-align: center;
            }}
            .wholepage .mark-tables td table tr .wordnumcolumn{{
                text-align: left;
                padding-left: 2px;
                width: 3.5cm;
            }}
            .wholepage .mark-tables td table tr .mocolumn{{
                padding-left: 2px;
                width: 1cm;
            }}
            .white {{
                color: white;
            }}
            .highlight {{
                font-weight: bold;
            }}
            .underline {{
                text-decoration: underline;
                text-decoration-color: black;
            }}
            .bold {{
                font-weight: bold;
            }}
            .center {{
                text-align: center;
            }}
            .italics {{
                font-style: italic;
            }}
            .mark-tables .mt1 {{
                padding-right: 5px;
                padding-left: 0.5cm;
                width: 5cm;
            }}
            .mark-tables .mt2, .mark-tables .mt3 {{
                padding-left: 5px;
            }}
            .mark-table1, .mark-table2, .mark-table3 {{
                border: 1px solid black;
                border-collapse: collapse;
            }}
            .mark-table1 .codecolumn {{
                width: 2cm;
            }}
            .mark-table2 .mocolumn, .mark-table3 .mocolumn {{
                padding: 2px;
                text-align: center;
            }}
            .mark-table2 .codecolumn, .mark-table3 .codecolumn {{
                padding-left: 5px;
                padding-right: 5px;
            }}
            .rows {{
                height:{19/number_of_rows}cm;
            }}
        </style>
        </head>
        <body>'''
    for i in range(pages):
        content1 = ""
        content2 = ""
        content1, row1_list, pass_counts, fail_counts = splice_pages(row1_list, pass_marks, number_of_rows) #splice first n off
        content2, row2_list, _ , _ = splice_pages(row2_list, None, number_of_rows) #splice the rest

        if thpr == "Pr.":
            combined_html_content += f'''
                <div id='overlay' style='
                    position: absolute;
                    top: 0.3cm;
                    left: 12.5cm;
                    z-index: 9999;
                    border: 1px solid;
                    padding: 1px;
                    text-align: center;
                '>
                    {campus.get()}
                </div>
            '''

        combined_html_content += f'''
        <table class="wholepage">
            <tr>
                <td class="col1">
                    <div class="center bold times">Tribhuvan University</div>
                </td>
                <td class="col23">
                    <div class="center bold times">Tribhuvan University</div>
                </td>
                <td class="col23 center">
                    <div class="bold times" style="display: inline-block;">Tribhuvan University</div>
                </td>
            </tr>
            <tr>
                <td class="col1">
                    <div class="center bold times">Institute of Medicine</div>
                </td>
                <td class="col23">
                    <div class="center bold times">Institute of Medicine</div>
                </td>
                <td class="col23">
                    <div class="center bold times">Institute of Medicine</div>
                </td>
            </tr>
            <tr class="mini">
                <td class="col1">
                    <div class="underline bold times">Counterfoil Examination</div>
                </td>
                <td class="col23">
                    <div class="underline bold italics times">To be treated as strictly confidential</div>
                </td>
                <td class="col23">
                    <div class="underline bold italics times">To be treated as strictly confidential</div>
                </td>
            </tr>
            <tr class="mini">
                <td class="col1">
                    <div>&nbsp;&nbsp;Examination held {data["date"]}</div>
                </td>
                <td class="col23">
                    <div>&nbsp;&nbsp;Examination held {data["date"]}</div>
                </td>
                <td class="col23">
                    <div>&nbsp;&nbsp;Examination held {data["date"]}</div>
                </td>
            </tr>
            <tr class="mini">
                <td class="col1">
                    <div>&nbsp;&nbsp;Level: {data["level"]}</div>
                </td>
                <td class="col23">
                    <div>&nbsp;&nbsp;Please follow directions strictly</div>
                </td>
                <td class="col23">
                    <div>&nbsp;&nbsp;Please follow directions strictly</div>
                </td>
            </tr>
            <tr class="mini">
                <td class="col1">
                    <div>&nbsp;&nbsp;Program: {data["program"]}&nbsp;&nbsp;Year: {data["year"]}  </div>
                </td>
                <td class="col23">
                    <div>&nbsp;&nbsp;Level: {data["level"]}&nbsp;&nbsp;Program: {data["program"]}</div>
                </td>
                <td class="col23">
                    <div>&nbsp;&nbsp;Level: {data["level"]}&nbsp;&nbsp;Program: {data["program"]}</div>
                </td>
            </tr>
            <tr class="mini">
                <td class="col1">
                    <div>&nbsp;&nbsp;Subject: {data["subject"]}</div>
                </td>
                <td class="col23">
                    <div>&nbsp;&nbsp;Year: {data["year"]}  Subject: {data["subject"]}</div>
                </td>
                <td class="col23">
                    <div>&nbsp;&nbsp;Year: {data["year"]}  Subject: {data["subject"]}</div>
                </td>
            </tr>
            <tr class="mini">
                <td class="col1">
                    <div>&nbsp;&nbsp;Paper: {data["paper"]}&nbsp;&nbsp;&nbsp;&nbsp;<span class="bold">{thpr}</span></div>
                </td>
                <td class="col23">
                    <div>&nbsp;&nbsp;Paper: {data["paper"]}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="bold">{thpr}</span></div>
                </td>
                <td class="col23">
                    <div>&nbsp;&nbsp;Paper: {data["paper"]}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="bold">{thpr}</span></div>
                </td>
            </tr>
            <tr class="mini">
                <td class="col1">
                    <div>&nbsp;&nbsp;Full Marks: {data["full_marks"]}&nbsp;&nbsp;Pass Marks: {pass_marks}</div>
                </td>
                <td class="col23">
                    <div>&nbsp;&nbsp;Full Marks: {data["full_marks"]}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Pass Marks: {pass_marks}</div>
                </td>
                <td class="col23">
                    <div>&nbsp;&nbsp;Full Marks: {data["full_marks"]}&nbsp;&nbsp;&nbsp;&nbsp;Pass Marks: {pass_marks}</div>
                </td>
            </tr>
            <tr class="mark-tables">
                <td class="mt1">
                    <table class="mark-table1">
                            <tr>
                                <th class="times huge codecolumn">Code #</th>
                                <th class="times huge mocolumn">Marks Obtained</th>
                            </tr>
                        {content1}
                    </table>
                </td>
                <td class="mt2">
                    <table class="mark-table2">
                        <tr>
                            <th class="times huge codecolumn">Code #</th>
                            <th class="times mocolumn">Marks Obtained</th>
                            <th class="times huge wordnumcolumn">Marks in Words</th>
                        </tr>
                        {content2}
                    </table>
                </td>
                <td class="mt3">
                    <table class="mark-table3">
                        <tr>
                            <th class="times huge codecolumn">Code #</th>
                            <th class="times mocolumn">Marks Obtained</th>
                            <th class="times huge wordnumcolumn">Marks in Words</th>
                        </tr>
                        {content2}
                    </table>
                </td>
            </tr>
            <tr class="after-marks mini">
                <td class="col1">
                    <div class="times">Passed: <span class="underline">{pass_counts}</span>&nbsp;&nbsp;Failed: <span class="underline">{fail_counts}</span></div>
                    <div style="height:0.3cm;"></div>
                    <div class="underline">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></div>
                    <div class="times">Full Signature of Examiner'''
        if thpr == "Pr.":
            combined_html_content += f'''s'''
        combined_html_content += f'''
            </div>
                    <div class="underline">Note:</span></div>
                    <div style="margin-left:20px; font-size: 8pt;">a) All corrections must be initialed. This counter-foil must be preserved by the Examiner as the case may be for six months.</div>
                </td>
                <td class="col23">
                    <div class="times center" style="padding-top:10px;">&nbsp;&nbsp;&nbsp;&nbsp;Passed: <span class="underline">{pass_counts}</span>&nbsp;&nbsp;&nbsp;&nbsp;Failed: <span class="underline">{fail_counts}</span></div>
                    '''
        if thpr == "Th.":
            combined_html_content += f'''
                    <div>&nbsp;&nbsp;&nbsp;&nbsp;<span>{data["name"]}</span></div>
                    <div class="times" style="border-top: 1px solid black;padding-top: 2px;">&nbsp;&nbsp;&nbsp;&nbsp;<span>Full <span class="bold">NAME</span> of Examiner</span></div>
                    '''
        else:
            combined_html_content += f'''
                    <div style="height:0.5cm;"><br></div>
                    <div class="times" style="border-top: 1px solid black;padding-top: 2px;">&nbsp;&nbsp;&nbsp;&nbsp;<span>Full Signature of Examiners</span></div>
                    '''
        combined_html_content += f'''
                    <div>&nbsp;&nbsp;&nbsp;&nbsp;<span class="underline">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></div>
                    <div>Full Signature of Scrutiny Board, if any</div>
                    <div><span class="underline">Note:</span></div>
                    <div style="margin-left:0.5cm; font-size: 8pt;">The Scrutiny Board shall Scrutinize and verify the marks submitted by the examiner before forwarding the same to the Examination Division.</div>
                </td>
                <td class="col23">
                    <div class="times center" style="padding-top:10px;">&nbsp;&nbsp;&nbsp;&nbsp;Passed: <span class="underline">{pass_counts}</span>&nbsp;&nbsp;&nbsp;&nbsp;Failed: <span class="underline">{fail_counts}</span></div>
                    <div style="height:0.3cm;"></div>
                    <div>&nbsp;&nbsp;&nbsp;&nbsp;<span class="underline">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></div>
                
                    <div class="times">&nbsp;&nbsp;&nbsp;&nbsp;Full 
                    '''
        if thpr == "Pr.":
            combined_html_content += f'''
                    Signature of Examiners</div>
                    '''
        else:
            combined_html_content += f'''
                    <span class="bold">Signature</span> of Examiner</span></div>'''
        combined_html_content += f'''
                    <div>&nbsp;&nbsp;&nbsp;&nbsp;<span class="underline">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></div>
                    <div>Full Signature of Scrutiny Board, if any</div>
                    <div><span class="underline">Note:</span></div>
                    <div style="margin-left:0.5cm; font-size: 8pt;">The Scrutiny Board shall Scrutinize and verify the marks submitted by the examiner before forwarding the same to the Examination Division.</div>
                </td>
            </tr>
        </table>'''
        if i != pages - 1:
            combined_html_content += '''
                <div style='page-break-before:always'></div>
            '''
    combined_html_content += f'''
            </body>
            </html>'''

    # Show file dialog to choose the PDF filename and folder
    root = tk.Tk()
    root.destroy()
    if switch == 1:
        pdf_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
    else:
        pdf_file="markentry_shifted.pdf"
    if not pdf_file: # If the user cancels the file dialog, return without generating the PDF
        return

    # Generate the PDF with the provided filename and folder
    generate_pdf(combined_html_content, pdf_file, switch)

def on_closing(): # save entered data to persistence file on app close
    data = {
        "date": date_entry.get(),
        "level": level_entry.get(),
        "program": program_entry.get(),
        "year": year_entry.get(),
        "subject": subject_entry.get(),
        "paper": paper_entry.get(),
        "full_marks": full_marks_entry.get(),
        "type_of_exam": exam_type_var.get(),
        "name": name_entry.get(),
        "num_of_rows": num_of_rows.get(),
        "campus": campus.get(),
        "text_data": text_entry.get("1.0", tk.END).strip(),
    }
    save_data(data)
    app.destroy()

if __name__ == "__main__": # app gui and relevant logic
    import_modules()
    pdf_created = 0
    default_rows = 30
    rownum = 0
    app = tk.Tk()
    app.title("Markfoil generator")
    style = ttk.Style()
    date_label = ttk.Label(app, text="Date:")
    date_label.grid(row=rownum, column=0)
    date_entry = ttk.Entry(app)
    date_entry.grid(row=rownum, column=1)

    rownum += 1

    level_label = ttk.Label(app, text="Level:")
    level_label.grid(row=rownum, column=0)
    level_entry = ttk.Entry(app)
    level_entry.grid(row=rownum, column=1)

    rownum += 1

    program_label = ttk.Label(app, text="Program:")
    program_label.grid(row=rownum, column=0)
    program_entry = ttk.Entry(app)
    program_entry.grid(row=rownum, column=1)

    rownum += 1

    year_label = ttk.Label(app, text="Year:")
    year_label.grid(row=rownum, column=0)
    year_entry = ttk.Entry(app)
    year_entry.grid(row=rownum, column=1)

    rownum += 1

    subject_label = ttk.Label(app, text="Subject:")
    subject_label.grid(row=rownum, column=0)
    subject_entry = ttk.Entry(app)
    subject_entry.grid(row=rownum, column=1)

    rownum += 1

    paper_label = ttk.Label(app, text="Paper:")
    paper_label.grid(row=rownum, column=0)
    paper_entry = ttk.Entry(app)
    paper_entry.grid(row=rownum, column=1)

    rownum += 1

    full_marks_label = ttk.Label(app, text="Full Marks:")
    full_marks_label.grid(row=rownum, column=0)
    full_marks_entry = ttk.Entry(app)
    full_marks_entry.grid(row=rownum, column=1)

    rownum += 1

    exam_type_label = ttk.Label(app, text="Type of Exam:")
    exam_type_label.grid(row=rownum, column=0)
    exam_type_var = tk.StringVar(value="Theory")
    exam_type_switch = ttk.Combobox(app, textvariable=exam_type_var, values=["Theory", "Practical"], state="readonly")
    exam_type_switch.grid(row=rownum, column=1)

    rownum += 1
    
    name_label = ttk.Label(app, text="Name of Examiner")
    name_label.grid(row=rownum, column=0)
    name_entry = ttk.Entry(app)
    name_entry.grid(row=rownum, column=1)

    rownum += 1
    
    campus_label = ttk.Label(app, text="Campus short name")
    campus_label.grid(row=rownum, column=0)
    campus = ttk.Entry(app)
    campus.grid(row=rownum, column=1)

    rownum += 1
    
    num_of_rows_label = ttk.Label(app, text="Number of rows per page")
    num_of_rows_label.grid(row=rownum, column=0)
    num_of_rows = ttk.Entry(app)
    num_of_rows.grid(row=rownum, column=1)

    rownum += 1
    
    text_label = ttk.Label(app, text="Enter marks as\nCode number <SPACE> Mark obtained\nCode number <TAB> Mark obtained\n<ENTER> to type next mark")
    text_label.grid(row=rownum, column=0, columnspan=2)

    rownum += 1

    text_entry = tk.Text(app, wrap="none", width=40, height=10)
    text_entry.grid(row=rownum, column=0, columnspan=2)

    rownum += 1
    
    generate_button = ttk.Button(app, text="Generate PDF")
    generate_button.grid(row=rownum, column=0, columnspan=2)

    def normal_click(event):
        submit_data(1)

    def shift_click(event):
        submit_data(2)

    generate_button.bind("<Button-1>", normal_click)
    generate_button.bind("<Shift-Button-1>", shift_click)

    # Load data if available
    data = load_data()
    date_entry.insert(tk.END, data.get("date", ""))
    level_entry.insert(tk.END, data.get("level", ""))
    program_entry.insert(tk.END, data.get("program", ""))
    year_entry.insert(tk.END, data.get("year", ""))
    subject_entry.insert(tk.END, data.get("subject", ""))
    paper_entry.insert(tk.END, data.get("paper", ""))
    full_marks_entry.insert(tk.END, data.get("full_marks", ""))
    exam_type_var.set(data.get("type_of_exam", "Theory"))
    name_entry.insert(tk.END, data.get("name", ""))
    num_of_rows.insert(tk.END, data.get("num_of_rows", ""))
    campus.insert(tk.END, data.get("campus", ""))
    text_entry.insert(tk.END, data.get("text_data", ""))

    def fields_disabler():
        if exam_type_switch.get() == "Theory":
            campus.grid_remove()
            campus_label.grid_remove()
            name_entry.grid()
            name_label.grid()
        else:
            campus.grid()
            campus_label.grid()
            name_entry.grid_remove()
            name_label.grid_remove()

    fields_disabler()
    
    def on_exam_type_select(event):
        fields_disabler()

    # Bind the dropdown selection event to the function
    exam_type_switch.bind("<<ComboboxSelected>>", on_exam_type_select)

    current_number = 1

    rownum += 1

    # Function to handle the "Enter" key press
    def on_enter(event):
        try:
            pass_marks = "{:g}".format(int(data["full_marks"]) / 2)
        except (KeyError, ValueError):
            result_label = ttk.Label(app, text="Enter a number in pass marks!")
            result_label.grid(row=14, column=0, columnspan=2)
        global current_number
        # error_label.config(text="")  # Clear previous error message
        marks = event.widget.get("insert linestart", "insert lineend")
        
        if marks.strip():
            # Using regular expression to split at a delimiter (tab, space, comma, colon, or hyphen)
            parts = re.split(r'[\t ,:\-]+', marks, maxsplit=1)
            
            # Checking if there are exactly 2 parts (number and mark)
            if len(parts) != 2:
                show_toast("Type code number <SPACE or TAB> \nand marks or description.")
                # error_label.config(text="Type code number <SPACE or TAB> and marks or description.")
                return "break"
            
            try:
                current_number_str, entered_mark_str = parts
                current_number = int(current_number_str) + 1

                # Checking if the current number is an integer before validation
                if current_number_str.isdigit():
                    entered_mark = float(entered_mark_str)  # Extract mark from tab-separated input
                    if entered_mark > float(full_marks_entry.get()):
                        # Display a message or take appropriate action
                        show_toast("Entered mark exceeds full marks")
                        # error_label.config(text="Entered mark exceeds full marks")
                        return "break"  # Prevent default newline behavior
            except ValueError:
                # Code to handle the case where the mark couldn't be extracted or converted
                pass

            delimiter_match = re.search(r'[\t ,:\-]+', marks)
            if delimiter_match:
                delimiter = delimiter_match.group()
                text_entry.insert(tk.END, f"\n{current_number}{delimiter}")
                text_entry.mark_set(tk.INSERT, f"insert +{len(str(current_number)) + len(delimiter)}c")
                text_entry.yview_moveto(1.0)  # Scroll to the new line

            return "break"  # Prevent default newline behavior

    # Bind the "Enter" key press event to the text entry
    text_entry.bind("<Return>", on_enter)

    # Bind the on_closing function to the window closing event
    app.protocol("WM_DELETE_WINDOW", on_closing)

    app.mainloop()
