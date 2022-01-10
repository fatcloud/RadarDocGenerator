from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from math import ceil
from PIL import Image
import os


def fill_cell(cell, txt):
    cell.paragraphs[0].runs[0].text = txt


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
        length = self.policy_length = len(self.shortbaseline)
        assert length in [39, 69]
        self.document = Document("templates\\template-" + str(length) + "e.docx")

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
        
        # 驗證範例檔的表格列數與資料的欄位數量相符 (照理說不會有問題)
        row_number = ceil(self.policy_length / 2)
        assert row_number == len(table.rows)-1

        # 開始填表 (左半)
        for idx in range(row_number):
            day1, day2, distance, period = self.shortbaseline[idx]
            cells = table.rows[idx+1].cells
            fill_cell(cells[1], day1 + '-' + day2)
            fill_cell(cells[2], str(round(float(distance), 3)))
            fill_cell(cells[3], period.strip())

        # 繼續填表 (右半)
        for idx in range(row_number, len(self.shortbaseline)):
            day1, day2, distance, period = self.shortbaseline[idx]
            cells = table.rows[idx - row_number + 1].cells
            fill_cell(cells[5], day1 + '-' + day2)
            fill_cell(cells[6], str(round(float(distance), 3)))
            fill_cell(cells[7], period.strip())

    def fill_image_table(self):
        table = self.document.tables[2]

        # 準備壓縮圖片的資料夾
        temppath = self.policy_dir + '\\temp-compressed'
        if not os.path.isdir(temppath):
            os.mkdir(temppath)

        # 填表
        cell_per_row = {39:5, 69:7}[self.policy_length]
        cell_size = {39:[2.9, 2.3], 69:[2.1, 1.7]}[self.policy_length]
        cell_scale = {39:1, 69:0.75}[self.policy_length]
        for idx, row in enumerate(self.shortbaseline):
            day1, day2, distance, period = self.shortbaseline[idx]

            # 填寫日期區段
            cell = table.rows[idx//cell_per_row * 2].cells[idx % cell_per_row]
            fill_cell(cell, day1 + '-' + day2)
            cell.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER

            # 壓縮圖片
            filename = day1 + '-' + day2 + '.tflt.filt.de'
            bmp_path = self.policy_dir + '\\postprocessing\\detrend_obs_file\\' + filename + '.bmp'
            png_path = temppath + '\\' + filename + '.png'
            img = Image.open(bmp_path)
            w, h = img.size
            s = cell_scale * self.image_scale
            img.resize((ceil(w*s),ceil(h*s))).save(png_path)

            # 填入圖片
            cell = table.rows[idx//cell_per_row * 2 + 1].cells[idx % cell_per_row]
            cell._element.clear_content()
            cell.add_paragraph().add_run().add_picture(png_path, height=Cm(cell_size[1]))
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
        self.document.save(self.policy_dir + '\\' + 'out.docx')



if __name__ == '__main__':

    import sys
    policy_path = sys.argv[1]
    print('=== 雷達影像基線資訊報表產生器 ===')
    print('啟動中...')
    print('正在 ' + policy_path + ' 位置下尋找 Policy 資料夾... ', end='')

    from docxcompose.composer import Composer
    from docx import Document as Document_compose

    master = None
    composer = None

    import glob
    policies = glob.glob(policy_path + '\Policy*')
    print('共找到 ', len(policies), ' 組資料如下')
    for p in policies:
        print(p.split('\\')[-1])

    print('開始處理:')

    for index, policy_dir in enumerate(policies):
        # 如果沒有 postprocessing 資料夾的話，這筆 Policy 會被跳過
        if os.path.isdir(policy_dir + '\\postprocessing'):
            print('processing', policy_dir, '...', end = '')

        # 創一個 Policy 物件
        doc_index = index+1
        policy = Policy(policy_dir=policy_dir, index=doc_index)
        
        # 輸出 docx，最後一頁不要換頁
        add_page_break = (doc_index != len(policies))
        policy.export_parital_docx(add_page_break=add_page_break)
        print('done')

        if master is None:
            master = Document_compose(policy_dir + '\\out.docx')
            composer = Composer(master)
        else:
            partial_docx = Document_compose(policy_dir + '\\out.docx')
            composer.append(partial_docx)

    composer.save(policy_path + "\output.docx")

