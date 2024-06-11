from docx import Document

# 워드 파일 열기
doc = Document("agenda/11-24-0964-05-00bn-may-july-tgbn-teleconference-agenda.docx")

submission_list = []
table = doc.tables[2]
for row in table.rows:
    year = row.cells[0].text.split("/")[0]
    if not year.isdigit():
        continue
    if len(year) > 2:
        dcn = year
        year = "2023"
    else:
        # print(submission_idx)
        # print(row.cells[0].text)
        # print(row.cells[0].text)
        dcn = row.cells[0].text.split("/")[1]
        year = "20" + year
    topic = row.cells[4].text
    session = row.cells[5].text
    items = [year, dcn, topic, session]
    submission_list.append(items)
print(submission_list)
