"""再計算（LibreOffice）で置き換わったフォント名を、xlsx内のstyles.xmlで指定フォントに戻す（計算結果は保持）
usage: python3 scripts/fix_font.py <file.xlsx> <font name>"""
import re, shutil, sys, tempfile, zipfile

path, font = sys.argv[1], sys.argv[2]
tmp = tempfile.mktemp(suffix='.xlsx')
with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'xl/styles.xml':
            text = data.decode('utf-8')
            text = re.sub(r'(<name val=")[^"]*(")', lambda m: m.group(1) + font + m.group(2), text)
            data = text.encode('utf-8')
        zout.writestr(item, data)
shutil.move(tmp, path)
