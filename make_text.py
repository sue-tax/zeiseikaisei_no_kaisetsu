'''
Created on 2024/10/14

@author: sue-t
'''

from pdfminer.high_level import extract_pages
from pdfminer.layout import LAParams, LTTextBoxHorizontal, LTRect


class ConvertKaisei(object):
    '''
    税制改正の解説のPDFファイルからテキストを抽出する
    '''


    def __init__(self, input_file_name, output_file_name):
        '''
        Constructor
        '''
        self.input_file = input_file_name
        self.output_text = '{}.txt'.format(output_file_name)
        self.output_md = '{}.md'.format(output_file_name)
        self.flag_text = True
        self.flag_md = True
        self.border = 261
        self.footer = 30
        self.header = 690
        self.start_page = 1 # 開始ページ1スタート
        self.last_page = 0  # 終了ページ

    def sorty_func(self, e):
        return - e.y1

    def text_in_rect(self, textbox, rect):
        if textbox.y1 > rect.y1 + 0.05:
            return False
        if textbox.y0 < rect.y0 - 0.05:
            return False
        if textbox.x0 < rect.x0 - 0.05:
            return False
        if textbox.x1 > rect.x1 + 0.05:
            return False
        return True

    def make_honbun(self, list_textbox, list_rect):
        if self.flag_start:
            start_text = 1
        else:
            start_text = 0
        left_honbun = []
        right_honbun = []
        for textbox in list_textbox[start_text:]:
            if len(list_rect) != 0:
                for rect in list_rect:
                    if self.text_in_rect(textbox, rect):
                        break
                else:
                    if textbox.x0 < self.border and \
                            textbox.x1 > self.border:
                        self.list_honbun.extend(left_honbun)
                        self.list_honbun.extend(right_honbun)
                        left_honbun = []
                        right_honbun = []
                        text = textbox.get_text(). \
                                replace(chr(0xfffd), '…')
                        self.list_honbun.append(text)
                    else:
                        if textbox.x1 < self.border:
                            text = textbox.get_text(). \
                                    replace(chr(0xfffd), '…')
                            left_honbun.append(text)
                        else:
                            text = textbox.get_text(). \
                                    replace(chr(0xfffd), '…')
                            right_honbun.append(text)
            else:
                if textbox.x0 < self.border and \
                        textbox.x1 > self.border:
                    self.list_honbun.extend(left_honbun)
                    self.list_honbun.extend(right_honbun)
                    left_honbun = []
                    right_honbun = []
                    text = textbox.get_text(). \
                            replace(chr(0xfffd), '…')
                    self.list_honbun.append(text)
                else:
                    if textbox.x1 < self.border:
                        text = textbox.get_text(). \
                                replace(chr(0xfffd), '…')
                        left_honbun.append(text)
                    else:
                        text = textbox.get_text(). \
                                replace(chr(0xfffd), '…')
                        right_honbun.append(text)
        self.list_honbun.extend(left_honbun)
        self.list_honbun.extend(right_honbun)

    def convert(self):
        laparams = LAParams()               # パラメータインスタンス
        laparams.boxes_flow = None          # -1.0（水平位置のみが重要）から+1.0（垂直位置のみが重要）default 0.5
        laparams.word_margin = 0.2          # default 0.1
        laparams.char_margin = 2.0          # default 2.0
        laparams.line_margin = 0            # default 0.5

        self.flag_start = True
        self.flag_mokuji = True
        
        self.list_mokuji = []
        self.list_honbun = []
        self.list_md = []
        for page_layout in extract_pages(self.input_file,
                maxpages=0, laparams=laparams):    # ファイルにwithしている
            list_textbox = []
            list_rect = []
            for element in page_layout:
                if isinstance(element, LTTextBoxHorizontal):
                    if element.y0 > self.header:
                        continue
                    if element.y1 < self.footer:
                        continue
                    list_textbox.append(element)
                if isinstance(element, LTRect):
                    list_rect.append(element)
            list_textbox.sort(key=self.sorty_func)
            # print(list_textbox)
            # for textbox in list_textbox:
            #     print(textbox)
            
            if self.flag_start:
                self.text_title = list_textbox[0].get_text().replace('\n','')
                self.list_md.append("# " + self.text_title + '\n\n')
                # print(self.text_title)
                self.text_mokuji = 2
                self.rect_mokuji = 3
                # self.flag_start = False
                # self.list_md.append("## 目次\n\n")
            if self.flag_mokuji:
                self.left_mokuji = []
                self.right_mokuji = []
                # print(page_layout.pageid)
                # print(self.rect_mokuji)
                self.rect_left = list_rect[self.rect_mokuji]
                self.rect_right = list_rect[self.rect_mokuji+2]
                # print(self.rect_left)
                # print(self.rect_right)
            
                for textbox in list_textbox[self.text_mokuji:]:
                    if self.text_in_rect(textbox, self.rect_left):
                        # text = textbox.get_text().replace('\n','')
                        # text = text.replace(chr(0x08), '…')
                        text = textbox.get_text()
                        # for ch in text:
                        #     print(hex(ord(ch)))
                        text = text.replace(chr(0xfffd), '…')
                        # print(text)
                        # for ch in text:
                        #     print(hex(ord(ch)))
                        self.left_mokuji.append(text)
                    if self.text_in_rect(textbox, self.rect_right):
                        # text = textbox.get_text().replace('\n','')
                        # text = text.replace(chr(0x08), '…')
                        text = textbox.get_text()
                        # for ch in text:
                        #     print(hex(ord(ch)))
                        text = text.replace(chr(0xfffd), '…')
                        # print(text)
                        self.right_mokuji.append(text)
                
                # print(self.left_mokuji)
                # for text in self.left_mokuji:
                #     print(text)
                # print(self.right_mokuji)
                # for text in self.right_mokuji:
                #     print(text)
                self.list_mokuji.extend(self.left_mokuji)
                self.list_mokuji.extend(self.right_mokuji)
                # list_without_rect = list_rect[self.rect_mokuji+2+1:]
                self.make_honbun(list_textbox, list_rect)
                        # self.rect_left, self.rect_right,
                        # list_without_rect)
                
                if self.rect_left.y0 > 55:
                    self.flag_mokuji = False
                else:
                    self.text_mokuji = 0
                    self.rect_mokuji = 0
            else:
                # list_without_rect = list_rect
                self.make_honbun(list_textbox, list_rect)
                        # None, None,
                        # list_without_rect)
            if self.flag_start:
                self.flag_start = False
                
            
            # break
        
        
        # for text in self.list_mokuji:
        #     print(text)
        # for text in self.list_honbun:
        #     print(text)
        if self.flag_text:
            f = open(self.output_text, 'w', encoding='UTF-8')
            f.write(self.text_title + '\n')
            # f.writelines(self.list_mokuji)
            for text in self.list_mokuji:
                f.write(text)
            f.writelines(self.list_honbun)
            f.close()

        if self.flag_md:
            import re
            p_dai = re.compile(
                    r'(第[一二三四五六七八九十]).?　([^\r\n\x08…]+)[^\r\n]*?(\r\n|\n|\r)')
            p_kansuji = re.compile(
                    r'([一二三四五六七八九十]+).?　([^\r\n\x08…]+)[^\r\n]*?(\r\n|\n|\r)')
            p_suji = re.compile(
                    r'([0-9０-９]+).?　([^\r\n\x08…]+)[^\r\n]*?(\r\n|\n|\r)')
            # ローマ数字１３、１４があるが、0x1a
            p_romasuji = re.compile(
                    r'([Ⅰ-Ⅻ]+).?　([^\r\n\x08…]+)[^\r\n]*?(\r\n|\n|\r)')
            p_kakko = re.compile(r'([⑴-⒇]) ?[ 　]([^\r\n]+?)(\r\n|\n|\r)')
            p_maru = re.compile(r'([①-⑳]) ?[ 　]([^\r\n]+?)(\r\n|\n|\r)')

            p_other = re.compile(
                    r'([^\r\n\x08…]+)[^\r\n]*?(\r\n|\n|\r)')
            
            self.list_md.append("[TOC]\n\n")
            self.list_md.append("## 目次\n\n")

            it = iter(self.list_mokuji)
            mokuji = next(it, None)
            while mokuji:
                # print(mokuji)
                # for ch in mokuji:
                #     print(hex(ord(ch)))
                m_dai = p_dai.match(mokuji)
                # print(m_dai)
                if m_dai:
                    text = m_dai.group(1) + '　' + m_dai.group(2)
                    # print(text)
                    mokuji = next(it, None)
                    while mokuji:
                        m_dai = p_dai.match(mokuji)
                        m_kansuji = p_kansuji.match(mokuji)
                        m_suji = p_suji.match(mokuji)
                        m_romasuji = p_romasuji.match(mokuji)
                        if m_dai or m_kansuji or m_suji or m_romasuji:
                            break
                        m_other = p_other.match(mokuji)
                        if not m_other:
                            break
                        text += m_other.group(1)
                        # print(text)
                        mokuji = next(it, None)
                    text = '[' + text + '](# ' + text + ')'
                    self.list_md.append(text + '\n')
                    continue
                m_kansuji = p_kansuji.match(mokuji)
                if m_kansuji:
                    text = m_kansuji.group(1) + '　' + m_kansuji.group(2)
                    # print(text)
                    mokuji = next(it, None)
                    while mokuji:
                        # print(mokuji)
                        # for ch in mokuji:
                        #     print(hex(ord(ch)))
                        m_dai = p_dai.match(mokuji)
                        m_kansuji = p_kansuji.match(mokuji)
                        m_suji = p_suji.match(mokuji)
                        m_romasuji = p_romasuji.match(mokuji)
                        if m_dai or m_kansuji or m_suji or m_romasuji:
                            break
                        m_other = p_other.match(mokuji)
                        if not m_other:
                            break
                        # print(m_other.groups())
                        text += m_other.group(1)
                        # print(text)
                        mokuji = next(it, None)
                    text = '[' + text + '](# ' + text + ')'
                    self.list_md.append(text + '\n')
                    continue
                m_suji = p_suji.match(mokuji)
                if m_suji:
                    text = m_suji.group(1) + '　' + m_suji.group(2)
                    # print(text)
                    mokuji = next(it, None)
                    while mokuji:
                        m_dai = p_dai.match(mokuji)
                        m_kansuji = p_kansuji.match(mokuji)
                        m_suji = p_suji.match(mokuji)
                        m_romasuji = p_romasuji.match(mokuji)
                        if m_dai or m_kansuji or m_suji or m_romasuji:
                            break
                        m_other = p_other.match(mokuji)
                        if not m_other:
                            break
                        # print(m_other.groups())
                        text += m_other.group(1)
                        # print(text)
                        mokuji = next(it, None)
                    text = '[' + text + '](# ' + text + ')'
                    self.list_md.append(text + '\n')
                    continue
                m_romasuji = p_romasuji.match(mokuji)
                if m_romasuji:
                    # text = m_romasuji.group(1) + '　' + m_romasuji.group(2)
                    text_title = m_romasuji.group(1) + '　' + m_romasuji.group(2)
                    text = '[' + text_title + '](# ' + text_title + ')'
                    # print(text)
                    mokuji = next(it, None)
                    while mokuji:
                        m_dai = p_dai.match(mokuji)
                        m_kansuji = p_kansuji.match(mokuji)
                        m_suji = p_suji.match(mokuji)
                        m_romasuji = p_romasuji.match(mokuji)
                        if m_dai or m_kansuji or m_suji or m_romasuji:
                            break
                        m_other = p_other.match(mokuji)
                        if not m_other:
                            break
                        # print(m_other.groups())
                        text += m_other.group(1)
                        # print(text)
                        mokuji = next(it, None)
                    self.list_md.append(text + '\n')
                    continue
                mokuji = next(it, None)
            self.list_md.append('\n')
            
            self.list_md.append('## 本文\n\n')
            it = iter(self.list_honbun)
            honbun = next(it, None)
            while honbun:
                # print(honbun)
                # for ch in mokuji:
                #     print(hex(ord(ch)))
                if honbun == 'はじめに\n':
                    self.list_md.append("#### " + honbun + "\n")
                    honbun = next(it, None)
                    continue
                m_dai = p_dai.match(honbun)
                # print(m_dai)
                if m_dai:
                    text = '## ' + m_dai.group(1) + '　' + m_dai.group(2)
                    # print(text)
                    honbun = next(it, None)
                    while honbun:
                        m_dai = p_dai.match(honbun)
                        m_kansuji = p_kansuji.match(honbun)
                        m_suji = p_suji.match(honbun)
                        m_romasuji = p_romasuji.match(honbun)
                        m_kakko = p_kakko.match(honbun)
                        m_maru = p_maru.match(honbun)
                        if m_dai or m_kansuji or m_suji or m_romasuji \
                                or m_kakko or m_maru:
                            break
                        m_other = p_other.match(honbun)
                        if not m_other:
                            break
                        text += m_other.group(1)
                        # print(text)
                        honbun = next(it, None)
                    self.list_md.append(text + '\n\n')
                    continue
                m_kansuji = p_kansuji.match(honbun)
                if m_kansuji:
                    text = '### ' + m_kansuji.group(1) + '　' + m_kansuji.group(2)
                    # print(text)
                    honbun = next(it, None)
                    while honbun:
                        m_dai = p_dai.match(honbun)
                        m_kansuji = p_kansuji.match(honbun)
                        m_suji = p_suji.match(honbun)
                        m_romasuji = p_romasuji.match(honbun)
                        m_kakko = p_kakko.match(honbun)
                        m_maru = p_maru.match(honbun)
                        if m_dai or m_kansuji or m_suji or m_romasuji \
                                or m_kakko or m_maru:
                            break
                        m_other = p_other.match(honbun)
                        if not m_other:
                            break
                        # print(m_other.groups())
                        text += m_other.group(1)
                        # print(text)
                        honbun = next(it, None)
                    self.list_md.append(text + '\n\n')
                    continue
                m_suji = p_suji.match(honbun)
                if m_suji:
                    text = '#### ' + m_suji.group(1) + '　' + m_suji.group(2)
                    # print(text)
                    honbun = next(it, None)
                    # while honbun:
                    #     m_dai = p_dai.match(honbun)
                    #     m_kansuji = p_kansuji.match(honbun)
                    #     m_suji = p_suji.match(honbun)
                    #     m_romasuji = p_romasuji.match(honbun)
                    #     m_kakko = p_kakko.match(honbun)
                    #     m_maru = p_maru.match(honbun)
                    #     if m_dai or m_kansuji or m_suji or m_romasuji \
                    #             or m_kakko or m_maru:
                    #         break
                    #     m_other = p_other.match(honbun)
                    #     if not m_other:
                    #         break
                    #     print(m_other.groups())
                    #     text += m_other.group(1)
                    #     print(text)
                    #     honbun = next(it, None)
                    self.list_md.append(text + '\n\n')
                    continue
                m_romasuji = p_romasuji.match(honbun)
                if m_romasuji:
                    text = '#### ' + m_romasuji.group(1) + '　' + m_romasuji.group(2)
                    # print(text)
                    honbun = next(it, None)
                    # while honbun:
                    #     m_dai = p_dai.match(honbun)
                    #     m_kansuji = p_kansuji.match(honbun)
                    #     m_suji = p_suji.match(honbun)
                    #     m_romasuji = p_romasuji.match(honbun)
                    #     m_kakko = p_kakko.match(honbun)
                    #     m_maru = p_maru.match(honbun)
                    #     if m_dai or m_kansuji or m_suji or m_romasuji \
                    #             or m_kakko or m_maru:
                    #         break
                    #     m_other = p_other.match(honbun)
                    #     if not m_other:
                    #         break
                    #     print(m_other.groups())
                    #     text += m_other.group(1)
                    #     print(text)
                    #     honbun = next(it, None)
                    self.list_md.append(text + '\n\n')
                    continue
                m_kakko = p_kakko.match(honbun)
                if m_kakko:
                    text = '##### ' + m_kakko.group(1) + '　' + m_kakko.group(2)
                    # print(text)
                    honbun = next(it, None)
                    self.list_md.append(text + '\n\n')
                    continue
                m_maru = p_maru.match(honbun)
                if m_maru:
                    text = '###### ' + m_maru.group(1) + '　' + m_maru.group(2)
                    # print(text)
                    honbun = next(it, None)
                    self.list_md.append(text + '\n\n')
                    continue
                self.list_md.append(honbun) # + '\n')
                honbun = next(it, None)
            f = open(self.output_md, 'w', encoding='UTF-8')
            f.writelines(self.list_md)
            f.close()
            

if __name__ == "__main__":
    cnv = ConvertKaisei("p0117-0316.pdf", "sample")
    cnv.convert()
        