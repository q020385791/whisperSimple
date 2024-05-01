# 關於本專案
本專案為whisper由語音轉文字並寫入Excel的範例紀錄

# default Model
https://huggingface.co/4bit/whisper-large-v2-ggml/tree/main
ggml-large-v2.bin

# 操作流程
1. 選擇音訊所在資料夾，本專案預設為使用.mp3檔案</br>
  如果需要使用其他檔案類型需要轉檔為whisper支援的檔案類型</br>
  修改以下這行改為您的檔案類型副檔名並作轉檔為whisper支援之檔案類型再做處理</br>

   ```C#
   var mp3Files = Directory.GetFiles(txtTargetFolder.Text, "*.mp3");
   ```
2.選擇對應csv資料夾(此為業務邏輯，可刪除)</br>

3.根據是否符合關鍵字可存至MatchFIle.xlsx與NotMatchFile.xlsx

# 一次只處理一個檔案
嘛，我曾經直接給他跑foreach，因為是非同步處理</br>
如果檔案一多起來記憶體是會爆炸的</br>
為了追求穩定性</br>
所以一次目前只處理一個檔案，一個一個來</br>
如果顯示卡夠強大想要多個非同步處理檔案</br>
請特別注意顯示卡能不能承受</br>
不然程式可是為跑到一半自動關閉的</br>
