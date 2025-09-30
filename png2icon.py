from PIL import Image

# PNGをICOに変換
image = Image.open("icon.png")
# 複数サイズを含むICOファイルとして保存
image.save("icon.ico", format='ICO', sizes=[(16,16), (32,32), (48,48), (64,64), (128,128), (256,256)])