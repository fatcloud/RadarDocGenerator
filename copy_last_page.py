import docx
import os
from copy import deepcopy

def duplicate_last_page_and_save(source_path, output_path, anchor_text):
    """
    打開一個源文件，找到錨點文字定義的“最後一頁”，
    將這一頁的內容複製並附加到文件末尾，然後另存為新文件。
    """
    if not os.path.exists(source_path):
        print(f"錯誤：找不到來源檔案 '{source_path}'")
        return

    try:
        # 載入源文件
        document = docx.Document(source_path)

        # 獲取所有頂層元素的 XML 表示
        all_body_elements = document.element.body[:]
        start_index = -1

        # 遍歷所有段落來尋找錨點文字
        for i, paragraph in enumerate(document.paragraphs):
            if anchor_text in paragraph.text:
                print(f"找到錨點文字：'{anchor_text}'")
                try:
                    p_element = paragraph._p
                    start_index = all_body_elements.index(p_element)
                    break
                except ValueError:
                    continue
        
        if start_index != -1:
            # 獲取從錨點開始到結尾的所有元素
            elements_to_duplicate = all_body_elements[start_index:]
            print(f"準備複製 {len(elements_to_duplicate)} 個元素...")
            
            # 將這些元素的深層複本附加到文件末尾
            # 在同一個文件中操作，deepcopy 對於帶有關聯的元素（如圖片）更安全
            for element in elements_to_duplicate:
                new_element = deepcopy(element)
                document.element.body.append(new_element)
            
            # 儲存被修改過的文件到一個新路徑
            document.save(output_path)
            print(f"成功複製最後一頁並另存為 '{output_path}'")
        else:
            print(f"錯誤：在文件中未找到指定的錨點文字 '{anchor_text}'。")

    except Exception as e:
        print(f"處理文件時發生錯誤：{e}")

# --- 執行 ---
source_file = os.path.join('templates', 'baseline.docx')
# 將輸出文件名更改為一個新的名字
output_file = 'baseline_duplicated.docx'
anchor = 'ALOS-2各期雷達影像對干涉圖'

duplicate_last_page_and_save(source_file, output_file, anchor)