import re
import PyPDF2
import arabic_reshaper
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

def fix_arabic(text: str) -> str:
    try:
        reshaped = arabic_reshaper.reshape(text)
        return reshaped
    except Exception:
        return text

def read_pdf(path: str) -> list:
    context = []
    with open(path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            text = page.extract_text()
            if text:
                context.extend(text.splitlines())
    return context

def filters(text: list[str]) -> dict:
    due_date: list = []
    end_of_payments: list = []
    amount: list = []
    row = {}

    for i in range(len(text)):
        if "Contract No" in text[i]:
            contract_no = text[i].split("Contract No")[1].split(":")[0].replace(". ", "")
            row["Contract No"] = contract_no

        if "Tenancy Start Date" in text[i]:
            tenancy_start_date = text[i].split("Tenancy Start Date")[1].split(":")[0]
            row["Tenancy Start Date"] = tenancy_start_date

        if "Tenancy End Date" in text[i]:
            tenancy_end_date = text[i].split("Tenancy End Date")[1].split(":")[0]
            row["Tenancy End Date"] = tenancy_end_date

        if "name/Founder" in text[i]:
            company_name = "".join(text[i:i + 2]).split("Organization")[0]
            company_name = company_name.rsplit(" ", 1)[0].replace("name/Founder", "")[:-3]
            row["Tenancy Name"] = fix_arabic(company_name)

        if "National Address" in text[i]:
            national_address = text[i].replace("National Address", "")
            row["National Address"] = national_address

        # 🔍 استخراج اسم المؤجر بدقة، مع إزالة كلمة "الاسم"
        if "Lessor Data" in text[i]:
            try:
                lessor_text = " ".join(text[i:i + 3])
                lessor_text = re.sub(r"\s+", " ", lessor_text).strip()
                print("\n🔍 Debug - lessor_text (raw):", repr(lessor_text))
                if "Name" in lessor_text:
                    name_part = lessor_text.split("Name", 1)[1].strip()
                    name_clean = re.sub(r"^[:\s]*الاسم[:\s]*", "", name_part).strip()
                    # قطع عند Nationality أو نقطتين
                    name_clean = re.split(r"Nationality|:", name_clean)[0].strip()

                    # حذف كل صورة لكلمة "الاسم" أو أي رموز مشوهة مشهورة في نهاية الاسم
                    name_clean = re.sub(r"(اﻻﺳﻢ|الاسم|\sالاسم|\sاﻻﺳﻢ)$", "", name_clean, flags=re.UNICODE).strip()

                    print("✅ Extracted Lessor Name (final):", repr(name_clean))
                    row["Lessor Name"] = fix_arabic(name_clean)
                else:
                    print("❌ Lessor Name not found in text")
                    row["Lessor Name"] = ""
            except Exception as e:
                print("❌ Error extracting Lessor Name:", e)
                row["Lessor Name"] = ""

        pattern = r"^\d+\.\d+\s+\d{4}-\d{2}-\d{2}\s+\d{4}-\d{2}-\d{2}.*\d{4}-\d{2}-\d{2}\s+\d{4}-\d{2}-\d{2}\s+\d+\s*$"
        payments = re.findall(pattern, text[i])
        if payments:
            payments = "".join(payments).split(" ")
            due_date.append(payments[5])
            end_of_payments.append(payments[4])
            amount.append(payments[0])

    row["Due Date"] = due_date[:]
    row["End of Payments"] = end_of_payments[:]
    row["Amount"] = amount[:]
    return row

def convert_to_excel(data, output_file: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Contract Data"

    headers = [
        "Contract No",
        "Tenancy Start Date",
        "Tenancy End Date",
        "Tenancy Name",
        "National Address",
        "Lessor Name",
        "Due Date",
        "End of Payments",
        "Amount"
    ]
    ws.append(headers)

    ws.append([
        data.get("Contract No", ""),
        data.get("Tenancy Start Date", ""),
        data.get("Tenancy End Date", ""),
        data.get("Tenancy Name", ""),
        data.get("National Address", ""),
        data.get("Lessor Name", ""),
        "", "", ""
    ])

    row_index = 2
    for due, end, amount in zip(data['Due Date'], data['End of Payments'], data['Amount']):
        ws.cell(row=row_index, column=7, value=due)
        ws.cell(row=row_index, column=8, value=end)
        ws.cell(row=row_index, column=9, value=amount)
        row_index += 1

    column_widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value:
                column_letter = get_column_letter(cell.column)
                cell_length = len(str(cell.value))
                if column_letter not in column_widths:
                    column_widths[column_letter] = cell_length
                else:
                    if cell_length > column_widths[column_letter]:
                        column_widths[column_letter] = cell_length

    for col_letter, col_width in column_widths.items():
        ws.column_dimensions[col_letter].width = col_width + 2

    wb.save(output_file)
    print(f"✅ Excel file saved successfully at: {output_file}")

# مسارات الملفات
pdf_path = r"C:\Users\ream8\Desktop\project\10988496532.pdf"
extracting = read_pdf(pdf_path)
pdf_data = filters(extracting)

excel_path = r"C:\Users\ream8\Desktop\project\Tenant_Info.xlsx"
convert_to_excel(pdf_data, excel_path)
