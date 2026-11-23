import docx
import shutil

src = 'output/20261123_吳修齊基金會_重溪國小感謝信_v4.docx'
dst = 'templates/重溪國小_正式信函公版_防裁切標準格式.docx'

shutil.copyfile(src, dst)
doc = docx.Document(dst)

# Replace paragraphs with template placeholders while preserving styles
placeholders = {
    0: '敬愛的 【受文單位董事長/長官尊稱】 鈞鑒：',
    1: '展信安好！',
    2: '    【請貼上 AI 生成第 1 段：由衷感謝、說明受核定計畫名稱與核定金額，表達全體師生振奮與溫暖】',
    3: '    【請貼上 AI 生成第 2 段：客觀描繪鄉村學校環境、學童與家庭背景，強調善款為孩子點亮學習明燈與免除家長後顧之憂】',
    4: '    【請貼上 AI 生成第 3 段：闡述教育大愛之深遠意義，承諾專款專用、深耕教學與陪伴孩子成長】',
    5: '    【請貼上 AI 生成第 4 段：說明檢附文件（如農會公庫領據、存摺封面影本、成果報告或致贈紀念品等）】',
    6: '耑此  敬祝',
    7: '平安喜樂  萬事如意',
    10: '中華民國 11X 年 XX 月 XX 日'
}

for idx, text in placeholders.items():
    if idx < len(doc.paragraphs):
        p = doc.paragraphs[idx]
        if p.runs:
            # Preserve first run font style if possible
            r0 = p.runs[0]
            # Clear other runs
            for r in p.runs[1:]:
                r.text = ''
            r0.text = text
        else:
            p.text = text

doc.save(dst)
print('Template successfully generated at:', dst)
