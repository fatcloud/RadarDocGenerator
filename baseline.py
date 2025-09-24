import pathlib
import utils
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from math import ceil
from PIL import Image
import os
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from copy import deepcopy

def set_table_borders(table):
    """
    Applies specific border styling to a table:
    - Thick outer borders
    - Thin inner borders
    """
    # Define border styles
    thin_border = {"sz": 4, "val": "single", "color": "auto"}
    thick_border = {"sz": 12, "val": "single", "color": "auto"}

    # Iterate through each cell to set borders
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')

            # Clear existing borders
            for border in tcPr.findall(qn('w:tcBorders')):
                tcPr.remove(border)

            # Determine border type based on cell position
            is_top = (i == 0)
            is_left = (j == 0)
            is_bottom = (i == len(table.rows) - 1)
            is_right = (j == len(row.cells) - 1)

            # Top border
            top_border = thick_border if is_top else thin_border
            border_el = OxmlElement('w:top')
            for key, val in top_border.items():
                border_el.set(qn(f'w:{key}'), str(val))
            tcBorders.append(border_el)

            # Left border
            left_border = thick_border if is_left else thin_border
            border_el = OxmlElement('w:left')
            for key, val in left_border.items():
                border_el.set(qn(f'w:{key}'), str(val))
            tcBorders.append(border_el)

            # Bottom border
            bottom_border = thick_border if is_bottom else thin_border
            border_el = OxmlElement('w:bottom')
            for key, val in bottom_border.items():
                border_el.set(qn(f'w:{key}'), str(val))
            tcBorders.append(border_el)

            # Right border
            right_border = thick_border if is_right else thin_border
            border_el = OxmlElement('w:right')
            for key, val in right_border.items():
                border_el.set(qn(f'w:{key}'), str(val))
            tcBorders.append(border_el)
            
            tcPr.append(tcBorders)

def fill_cell(cell, txt):
    # This function seems to assume there is a run already.
    # A safer version would be:
    if cell.paragraphs and cell.paragraphs[0].runs:
        cell.paragraphs[0].runs[0].text = txt
    else:
        cell.text = txt

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None


class Policy:

    def __init__(self, policy_dir, index, image_scale=0.15, areas=None):
        self.index = index
        self.areas = areas
        self.image_scale = image_scale
        self.policy_dir = policy_dir

        # 把 baseline 開出來
        self.shortbaseline = []
        with open(policy_dir + '\\postprocessing\\shortbaseline') as f:
            for line in f:
                self.shortbaseline.append(line.split('\t'))

        # 開啟對應的範例
        self.policy_length = len(self.shortbaseline)
        try:
            self.document = Document("templates\\baseline.docx")
        except Exception as e:
            raise FileNotFoundError("找不到 'templates\\baseline.docx' 範本檔案，請從 'templates' 資料夾中任選一個 .docx 檔案，並將其改名為 'baseline.docx'") from e

    def _get_layout_params(self):
        supported_lengths = [7, 27, 39, 69]
        # 找到最接近的支援長度
        closest_length = min(supported_lengths, key=lambda x: abs(x - self.policy_length))

        # 如果長度遠大於已知的最大長度，就使用最大長度的參數
        if self.policy_length > max(supported_lengths):
            closest_length = max(supported_lengths)

        cell_per_row_map = {39: 5, 69: 7, 27: 4, 7: 4}
        cell_size_map = {39: [2.9, 2.3], 69: [2.1, 1.7], 27: [3.6, 2.1], 7: [3.6, 3.6]}
        cell_scale_map = {39: 1, 69: 0.75, 27: 1, 7: 1}

        return (cell_per_row_map.get(closest_length, 4),
                cell_size_map.get(closest_length, [3.6, 3.6]),
                cell_scale_map.get(closest_length, 1))

    def check_bmp_path(self):
        for idx, row in enumerate(self.shortbaseline):
            day1, day2, distance, period = self.shortbaseline[idx]

            # 檢查檔案名
            filename = day1 + '-' + day2 + '.tflt.filt.de'
            bmp_path = self.policy_dir + '\\postprocessing\\detrend_obs_file\\' + filename + '.bmp'
            if not os.path.isfile(bmp_path):
                return bmp_path


    def export_eps_to_png(self):

        plot = 'shortbaseline_plot.eps'
        eps_image = Image.open( self.policy_dir + '\\' + plot)
        eps_image.load(scale=3)
        self.image_path = self.policy_dir + '\\' + 'shortbaseline_plot.png'
        eps_image.save(self.image_path)

    def extract_dates(self):
        days = []
        for data_row in self.shortbaseline:
            day1, day2, distance, period = data_row
            if day1 not in days:
                days.append(day1)
            if day2 not in days:
                days.append(day2)
        days.sort()
        return days
   
    def fill_dates(self):
        cell = self.document.tables[0].rows[1].cells[1]
        dates = self.extract_dates()
        fill_cell(cell, '、'.join(dates) + '。')

    def fill_index(self):
        index_run = self.document.paragraphs[2].runs[0]
        index_run.text = '(' + str(self.index) + ')'

    def fill_areas(self):
        # 預設是生成文件後再由人手動填入
        area_run = self.document.paragraphs[3].runs[0]
        areas = self.areas
        if not areas:
            return

        for i in range(2, len(areas), 2):
            areas[i] = '\n' + areas[i]
        area_run.text = '、'.join(areas) + '。'

    def fill_image_metadata(self):
        table = self.document.tables[1]
    
        row_number = ceil(self.policy_length / 2)

        # 動態增刪表格列
        current_rows = len(table.rows) - 1
        
        template_row = None
        if current_rows > 0 and len(table.rows) > 1:
            template_row = table.rows[1]

        if row_number > current_rows:
            for _ in range(row_number - current_rows):
                new_row = table.add_row()
                if template_row:
                    # 套用列高
                    new_row.height = template_row.height
                    new_row.height_rule = template_row.height_rule
                    # 複製儲存格格式 (包含底色)
                    for i, template_cell in enumerate(template_row.cells):
                        new_cell = new_row.cells[i]
                        new_tcPr = deepcopy(template_cell._tc.get_or_add_tcPr())
                        p = new_cell._tc.get_or_add_tcPr()
                        p.getparent().replace(p, new_tcPr)

        elif row_number < current_rows:
            for i in range(current_rows - row_number):
                row_to_remove = table.rows[current_rows - i]
                row_to_remove._element.getparent().remove(row_to_remove._element)

        # 清空可能存在的舊資料
        for r_idx in range(1, len(table.rows)):
            for c_idx in range(len(table.columns)):
                fill_cell(table.rows[r_idx].cells[c_idx], "")

        # 開始填表 (左半)
        for idx in range(row_number):
            if idx >= len(self.shortbaseline): break
            day1, day2, distance, period = self.shortbaseline[idx]
            cells = table.rows[idx+1].cells
            fill_cell(cells[0], str(idx + 1)) # 填寫 NO
            fill_cell(cells[1], day1 + '-' + day2)
            fill_cell(cells[2], str(round(float(distance), 3)))
            fill_cell(cells[3], period.strip())

        # 繼續填表 (右半)
        for idx in range(row_number, len(self.shortbaseline)):
            day1, day2, distance, period = self.shortbaseline[idx]
            row_index_in_table = idx - row_number + 1
            cells = table.rows[row_index_in_table].cells
            fill_cell(cells[4], str(idx + 1)) # 填寫 NO
            fill_cell(cells[5], day1 + '-' + day2)
            fill_cell(cells[6], str(round(float(distance), 3)))
            fill_cell(cells[7], period.strip())

        # 將表格內所有文字置中
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 套用邊框樣式
        set_table_borders(table)

    def fill_image_table(self):
        # 取得排版參數
        cell_per_row, cell_size, cell_scale = self._get_layout_params()

        # 移除舊表格並建立新表格
        old_table = self.document.tables[2]
        old_table._element.getparent().remove(old_table._element)
        
        num_rows = ceil(self.policy_length / cell_per_row) * 2
        table = self.document.add_table(rows=num_rows, cols=cell_per_row)
        table.style = 'Table Grid'

        # 準備壓縮圖片的資料夾
        tmp_path = self.policy_dir + '\\tmp\\compressed'
        pathlib.Path(tmp_path).mkdir(parents=True, exist_ok=True)

        # 填表
        for idx, row in enumerate(self.shortbaseline):
            day1, day2, distance, period = self.shortbaseline[idx]

            # 填寫日期區段
            cell = table.rows[idx//cell_per_row * 2].cells[idx % cell_per_row]
            fill_cell(cell, day1 + '-' + day2)
            cell.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER

            # 壓縮圖片
            filename = day1 + '-' + day2 + '.tflt.filt.de'
            bmp_path = self.policy_dir + '\\postprocessing\\detrend_obs_file\\' + filename + '.bmp'
            png_path = tmp_path + '\\' + filename + '.png'
            img = Image.open(bmp_path)
            w, h = img.size
            s = cell_scale * self.image_scale
            img.resize((ceil(w*s),ceil(h*s))).save(png_path)

            # 填入圖片
            cell = table.rows[idx//cell_per_row * 2 + 1].cells[idx % cell_per_row]
            cell._element.clear_content()
            if not cell.paragraphs:
                cell.add_paragraph()
            cell.paragraphs[0].add_run().add_picture(png_path, height=Cm(cell_size[1]))
            cell.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER

    def add_plot(self):
        image_paragraph = self.document.paragraphs[4]
        image_paragraph.clear()
        image_paragraph.add_run().add_picture(self.image_path, width=Cm(15.53))

    def export_parital_docx(self, add_page_break=True):
        self.fill_dates()
        self.fill_index()
        self.fill_areas()
        self.fill_image_metadata()
        self.fill_image_table()
        self.export_eps_to_png()
        self.add_plot()
        if add_page_break:
            self.document.add_page_break()
            
        doc_path = self.policy_dir + '\\tmp\\baseline-' + str(self.index) + '.docx'
        self.document.save(doc_path)

        return doc_path



if __name__ == '__main__':

    import sys
    policy_path = sys.argv[1]
    print('=== 雷達影像基線資訊報表產生器 ===')
    print('啟動中...')
    print('正在 ' + policy_path + ' 位置下尋找 Policy 資料夾... ', end='')

    import glob
    policies = glob.glob(policy_path + '\\Policy*')
    print('共找到 ', len(policies), ' 組資料如下')
    for p in policies:
        print(p.split('\\')[-1])

    print('開始處理:')
    doc_paths = []
    for index, policy_dir in enumerate(policies):
        # 如果沒有 postprocessing 資料夾的話，這筆 Policy 會被跳過
        if os.path.isdir(policy_dir + '\\postprocessing'):
            print('processing', policy_dir, '...', end = '')

        # 創一個 Policy 物件
        doc_index = index+1
        policy = Policy(policy_dir=policy_dir, index=doc_index)
        
        # 輸出 docx，最後一頁不要換頁
        add_page_break = (doc_index != len(policies))
        doc_path = policy.export_parital_docx(add_page_break=add_page_break)
        print('done')
        doc_paths.append(doc_path)

    # 組合所有頁面並打開檔案
    output_path = policy_path + "\\doc\\基線.docx"
    utils.concatenate_docx(doc_paths, output_path)
    os.system('start ' + output_path)
