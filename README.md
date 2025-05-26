# 📊 Excel Splitter Pro

**Excel Splitter Pro** is a sleek and user-friendly desktop application built with Python and Tkinter that allows you to split large Excel files into smaller CSV chunks effortlessly. Ideal for handling massive datasets, this tool simplifies data management by breaking down spreadsheets into manageable parts with just a few clicks.

---

## 🚀 Features

- **Easy Excel file selection** with a graphical file browser.
- **Choose output folder** to save split CSV files.
- Define **custom base filename** for all output chunks.
- Set **chunk size** to specify how many rows per CSV file.
- **Responsive progress bar and percentage indicator** during the splitting process.
- Clean success message with buttons to **open output folder** or **split another file**.
- Runs on Python with no complex dependencies besides `pandas` and `openpyxl`.
- Modern and intuitive GUI styled with `ttk` widgets for a polished user experience.

---

## 🎯 Why use Excel Splitter Pro?

Handling large Excel files can be cumbersome, especially when you need to work with smaller subsets of the data or upload to systems with size limits. This tool automates the splitting process without needing command-line knowledge or complicated scripts.

---

## 🛠️ Installation

1. **Clone this repository:**

```bash
git clone https://github.com/yourusername/excel-splitter-pro.git
cd excel-splitter-pro
````

2. **Install required Python packages:**

```bash
pip install pandas openpyxl
```

3. **Run the application:**

```bash
python excel_splitter_pro.py
```

---

## 🖥️ How to Use

1. Click **📂 Browse** next to **Excel File** to select your `.xlsx` or `.xls` file.
2. Click **📂 Browse** next to **Output Folder** to choose where the split CSV files will be saved.
3. Enter a base name for the output files (e.g., `data_chunk`).
4. Set the **Chunk Size** (number of rows per CSV file).
5. Click **⚡ Split Excel File** to start the process.
6. Watch the progress bar and percentage update in real-time.
7. Upon completion, a success message appears with options to:

   * **📁 Open Output Folder**
   * **🔄 Split Another File**

---

## 📂 Output

Split CSV files are saved in the chosen output folder with names formatted like:

```
{BaseName}_P1.csv
{BaseName}_P2.csv
...
```

---

## 🧑‍💻 Technologies Used

* Python 3.x
* Tkinter (GUI)
* pandas (Excel processing)
* openpyxl (Excel file reading engine)

---

## 📸 Screenshots

![Main Interface](screenshots/main_interface.png)
*Select your file, output folder, and chunk size.*

![Splitting Progress](screenshots/progress.png)
*Progress bar and percentage during splitting.*

![Success Screen](screenshots/success.png)
*Completion message with action buttons.*

---

## 💡 Contribution

Feel free to fork this project and submit pull requests!
If you find bugs or have suggestions, please open an issue.

---

## ⚖️ License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## 🙏 Acknowledgments

* Thanks to the creators of `pandas` and `openpyxl` for simplifying Excel data handling.
* Inspired by the need for a simple, GUI-based Excel splitter tool.

---

Happy splitting! 🎉
If you find this tool useful, ⭐ star the repo!

---

*Made with ❤️ using Python & Tkinter*
