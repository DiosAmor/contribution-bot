from docx import Document

# 워드 파일 열기
doc = Document("11-24-0964-01-00bn-may-july-tgbn-teleconference-agenda.docx")

# print(doc.paragraphs)
for table in doc.tables[2:3]:
    print(type(table))
    for row in table.rows[3:8]:
        for cell in row.cells:
            print(cell.text)

# 모든 단락 읽기
# for para in doc.paragraphs:
#     print(para.text)

# # 테이블 데이터 읽기
# for table in doc.tables:
#     for row in table.rows:
#         for cell in row.cells:
#             print(cell.text)
