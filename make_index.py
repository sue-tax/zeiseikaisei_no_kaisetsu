"""
税制改正の解説の目次用のExcelを作成する

途中

make_index2 に開発を移行

https://github.com/juu7g/Python-PDF2text
pdf_PDF2text
を参考にしている
"""
from pickle import NONE

'''
<LTRect 219.809,591.605,296.152,610.370>
<LTRect 46.233,591.604,219.809,600.987>
<LTRect 296.153,591.604,469.729,600.987>
<LTRect 46.233,427.366,248.768,591.605> 目次の左側
<LTRect 248.768,427.366,267.193,591.605>
<LTRect 267.194,427.366,469.729,591.605> 目次の右側

<LTTextBoxHorizontal(-1) 267.166,576.427,442.205,585.639 '五\u3000国庫補助金等の総収入金額不算入制度\n'>
<LTTextBoxHorizontal(-1) 55.414,576.408,230.454,585.621 '一\u3000新たな公益信託制度の創設に伴う所得\n'>
<LTTextBoxHorizontal(-1) 276.379,560.839,460.492,570.051 'の改正\x08��������������� 108\n'>
'''

from pdfminer.high_level import extract_pages
from pdfminer.layout import LAParams, LTTextBox, LTRect
import datetime
import collections
# import os, sys, argparse
import sys

import re

class ConvertPDF2textV():
    """
    PDFをtxtに変換する。
    PDFは2段組みの場合も含める
    """

    def __init__(self, argv:list):
        """
        コンストラクタ

         """
         
        # TODO フォルダ名　和暦、西暦の参考
        self.input_path = r'p0089-0116.pdf'
        self.output_path = '{}.txt'.format(datetime.datetime.now().strftime("%m%d_%H%M_%S"))
        self.border = 261     # 段組みの切れ目のx座標
        self.footer = 60    # フッターのy座標。ページの最下部が0。これより下の位置の文字は抽出しない
        self.header = 680  # ヘッダーのy座標。 これより上の位置の文字は抽出しない
        self.start_page = 1 # 開始ページ1スタート
        self.last_page = 0  # 終了ページ

        self.dict_tax = { '所得税法等の改正' : ('所得税', '所得税法'),
                         '租税特別措置法等（所得税関係）の改正' : ('所得税', '租税特別措置法')
                         }
        self.str_title = ''
        self.num_page = 0
        self.list_index = []
        self.dir_name = '.\\R06_2024'   # 暫定
        self.str_wareki = '令和6年'    # 暫定
        self.str_seireki = '2024年'    # 暫定

        self.p_index = re.compile('[一二三四五六七八九十]+.+?…+? [0-9]+')
        self.p_index1 = re.compile('[一二三四五六七八九十]+.+?')
        self.p_index2 = re.compile('.+?…+? [0-9]+')

        # self.input_path = args.input_path
        # self.output_path = args.output_path
        # self.sheet_name = os.path.splitext(os.path.basename(self.output_path))[0]

    def make_title(self, str_title):
        if not str_title.endswith("の改正"):
            return None
        tuple_title = self.dict_tax.get(str_title, None)
        return tuple_title
    
    def flatten(self, l):
        """
        ツリー状になっているイテレータをフラットに返すイテレータ
        """
        for el in l:
            if isinstance(el, collections.abc.Iterable) and not isinstance(el, (str, bytes)):
                yield from self.flatten(el)
            else:
                yield el

    def flatten_lttext(self, l, _type):
        """
        ツリー状になっているイテレータをフラットに返すイテレータ
        返る要素の型を指定
        pdfminerのextract_pagesで使用するのを想定
        要素の型が引数で指定した型を継承したもののみを返す

        Args:
            l:      pdfminerのextract_pages()の戻り値
            _type:  戻したい値の型
        """
        for el in l:
            if isinstance(el, (_type)):
                yield el
            else:
                if isinstance(el, collections.abc.Iterable) and not isinstance(el, (str, bytes)):
                    yield from self.flatten_lttext(el, _type)
                else:
                    continue

    def write2text(self, f):
        """
        ファイルへtext_l, text_rをこの順に書き込む
        書き込み後、text_l, text_rをクリア
        Args:
            f:      書き込みファイル
        """
        text = self.text_l.replace(chr(0x8), ' ')
        text = text.replace(chr(0xfffd), '…')
        
        # TODO パターンマッチングで目次内容かを確認
        list_text = text.split('\n')
        for line_text in list_text:
            m = self.p_index.match(line_text)
            if m != None:
                print(m)
                print(line_text)
            else:
                m1 = self.p_index1.match(line_text)
                if m1 == None:
                    continue
                print(line_text)
                line_text2 = list_text[list_text.index(line_text)+1]
                print(line_text2)
                m2 = self.p_index2.match(line_text2)
                print(m2)

        f.write(text)
        
        # for ch in self.text_l:
        #     print(ch, hex(ord(ch)), ord(ch))
        text = self.text_r.replace(chr(0x8), ' ')
        text = text.replace(chr(0xfffd), '…')

        # TODO パターンマッチングで目次内容かを確認
        list_text = text.split('\n')
        for line_text in list_text:
            m = self.p_index.match(line_text)
            if m != None:
                print(m)
                print(line_text)
            else:
                m1 = self.p_index1.match(line_text)
                if m1 == None:
                    continue
                print(line_text)
                line_text2 = list_text[list_text.index(line_text)+1]
                print(line_text2)
                m2 = self.p_index2.match(line_text2)
                print(m2)
        
        f.write(text)
        # for ch in self.text_r:
        #     print(ch, hex(ord(ch)), ord(ch))
        self.text_l = self.text_r = ""
    
    def write_footer(self, f, text_footer):
        """
        Args:
            f:      書き込みファイル
        """
        print(text_footer)
        print(text_footer[2:-3])
        f.write("###### 【ページ：")
        f.write(text_footer[2:-3])
        f.write("】\n")
        if self.num_page == 0:
            self.num_page = int(text_footer[2:-3])
            print("num_page ", self.num_page)

    def convert_pdf_to_text(self):
        """
        PDFファイルをテキストに変換
        PDFは2段に段組みされたものも含む
        """

        laparams = LAParams()               # パラメータインスタンス
        laparams.boxes_flow = None          # -1.0（水平位置のみが重要）から+1.0（垂直位置のみが重要）default 0.5
        laparams.word_margin = 0.2          # default 0.1
        laparams.char_margin = 2.0          # default 2.0
        laparams.line_margin = 0            # default 0.5

        # 出力ファイルのオープン    ファイルがある時は上書きされる
        with open(self.output_path, "w", encoding="utf-8") as f:
            # 初期化
            self.text_l = ""        # 左側の文字列
            self.text_r = ""        # 右側の文字列
            
            print("Analyzing from {} page to {} page(0:to last)".format(self.start_page, self.last_page))
            
            # todo 最初の文字列を保存　所得税法等の改正
            
            # todo 最初のページ数を保存
            
            # 対象ページを読み、テキスト抽出する。（maxpages：0は全ページ）
            for page_layout in extract_pages(self.input_path, maxpages=0, laparams=laparams):    # ファイルにwithしている
                
                # 抽出するページの選別。extract_pagesの引数では、開始ページだけの指定に対応できないため
                # if page_layout.pageid < self.start_page: continue                   # 指定開始ページより前は飛ばす
                # if self.last_page and self.last_page < page_layout.pageid: break    # 指定終了ページ以降は中断
                # ページの幅から段組みの境界を計算(用紙幅の半分とする)
                # if self.border == 0:
                #     self.border = int(page_layout.width / 2)
                if page_layout.pageid == self.start_page:
                    print("Check on page #{}".format(page_layout.pageid))
                    print("Page Info width:{}, heght:{}".format(page_layout.width, page_layout.height))
                    print("Calc result border: {}, footer: {}".format(self.border, self.footer))
                # 要素の出現順の確認(debug)
                # for element in self.flatten_lttext(page_layout, LTTextBox):
                #     print("bbox{} {}".format(element.bbox, element.get_text()[:20]))
                
                # print(page_layout)
                # for item in page_layout:
                #     print(item)
                #     if isinstance(item, LTRect):
                #         pass
                
                
                
                # 要素のイテレータをたどり入れ子の要素を1次元に取り出す。戻るイテレータはLTTextBox型のみ
                # 要素の行の上側y1で降順、行の左側x0で昇順にソートする。
                for element in sorted(self.flatten_lttext(page_layout, LTTextBox), key=lambda x: (-x.y1, x.x0)):
                # for element in self.flatten_lttext(page_layout, LTTextBox):
                    if element.y1 < self.footer:
                        _text =element.get_text()
                        self.write_footer(f, _text)
                        continue  # フッター位置の文字は抽出しない
                    if element.y0 > self.header: continue  # ヘッダー位置の文字は抽出しない
                    _text =element.get_text()
                    if self.str_title == '':
                        self.str_title = _text
                        print("str_title ", self.str_title)
                        self.tuple_title = self.make_title(self.str_title)
                    # debug
                    # print("y1:{}, y0:{}■{}".format(element.y1, element.y0, _text))

                    if element.x1 < self.border:
                        # 文字列全体が左側
                        self.text_l += _text
                    else:
                        if element.x0 >= self.border:
                            # 文字列全体が右側
                            self.text_r += _text
                        else:
                            # 文字列が境界をまたいでいる場合
                            # 右側に既に文章があれば先に出力する
                            if self.text_r:
                                self.write2text(f)
                            self.text_l += _text

                # 1ページ分処理したら書き込む
                self.write2text(f)

if __name__ == "__main__":
    cnv = ConvertPDF2textV(sys.argv[1:])
    cnv.convert_pdf_to_text()
