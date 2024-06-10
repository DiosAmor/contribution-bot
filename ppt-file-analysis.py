from pptx import Presentation

# PowerPoint 파일 열기
prs = Presentation("11-24-0653-15-00bn-tgbn-may-2024-meeting-agenda.pptx")

# 모든 슬라이드에서 텍스트 읽기
print(prs.slides[0])

# Slicing이 안 되는데 뭐지?
# getitem 쪽은 더 파봐야하는듯? rId가 뭐여
for slide in prs.slides[2:3]:
    print(type(slide))
    # for shape in slide.shapes:
    #     if hasattr(shape, "text"):
    #         print(shape.text)
