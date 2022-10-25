import aspose.words as aw
from datetime import date, datetime
import os
import io
import glob

def point_to_cm(pnt):
    value = aw.ConvertUtil.point_to_inch(pnt)*2.54
    return round(value, 1)

def cm_to_point(cm):
    value = aw.ConvertUtil.millimeter_to_point(cm*10)
    return round(value, 1)


class TextParagraph:
    
    def __init__(self) -> None:
        self.text = ''
        self.comment_ids = []
        self.aspose_paragraph_node = None
        
    def __str__(self) -> str:
        return "text: {0}, \ncomment_ids: {1}".format(self.text, self.comment_ids)
        
class TextComment:
    
    def __init__(self) -> None:
        self.text = None
        self.comment_id = None
        self.aspose_comment_node = None
        
    def __str__(self) -> str:
        return "text: {0}, \ncomment_id: {1}".format(self.text, self.comment_ids) 


class AsposeManager():    
    
    def __init__(self) -> None:
        self._set_license()
        self.doc = None
        self.origin_doc = None
        self.file_name = ''
        self.file_ext = 'docx'
        self.convert_ext = 'docx'
        self.convert_file_name = ''
        self.page_setup = {}
        
    def _set_license(self):
        lic = aw.License()
        try :
            lic.set_license("Aspose.Words.Python.NET.lic")
            print("License set successfully.")
        except RuntimeError as err :
            # We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license. 
            print("\nThere was an error setting the license: {0}".format(err))

    def read_document_stream(self, file):
        filename, fileExtension = os.path.splitext(file)
        self.file_name = filename
        stream = io.FileIO(file)
        self.doc = aw.Document(stream)
        stream.close()
        
    def read_document(self, file):
        
        filename, file_extension = os.path.splitext(file)
        self.file_name = filename
        self.file_ext = file_extension
        self.doc = aw.Document(file)
        
        if self.file_ext == '.html':
            self.origin_doc = aw.Document(self.file_name+".docx")
            builder = aw.DocumentBuilder(self.origin_doc)
            self.read_page_setup(builder.page_setup)
            
            for chunk in self.origin_doc.custom_document_properties:
                
                print(chunk.name)
                print(chunk.value)
            
        
        # print(dir(self.doc))
        # print(dir(self.doc.custom_document_properties))
        # print(self.doc.custom_document_properties.get_by_name('cprice'))
        # self.doc.custom_document_properties.remove('cprice')
        # self.doc.custom_document_properties.add('cprice', 458465)
        # self.doc.update_fields()
        
        
        
        
    def read_page_setup(self, page_setup):

        options = {}
        options['paper_size'] = aw.PaperSize(page_setup.paper_size).name
        options['top_margin'] = point_to_cm(page_setup.top_margin)
        options['bottom_margin'] = point_to_cm(page_setup.bottom_margin)
        options['left_margin'] = point_to_cm(page_setup.left_margin)
        options['right_margin'] = point_to_cm(page_setup.right_margin)
        options['header_distance'] = point_to_cm(page_setup.header_distance)
        options['footer_distance'] = point_to_cm(page_setup.footer_distance)

        self.page_setup = options
        
    def to_docx(self):
        self.convert_ext = 'docx'
        
    # todo: txt 변환시 docx 불필요 항목 삭제 방법 확인(예: 페이지 번호, 머리글, 꼬리글 등)
    def to_txt(self):
        comments = self.doc.get_child_nodes(aw.NodeType.COMMENT, True)
        comments.clear()
        
        for section in self.doc:
            section = section.as_section()
            section.headers_footers.clear()
            
        self.convert_ext = 'txt'
    
    def to_html(self):
        self.convert_ext = 'html'
        
    def to_pdf(self, comment_use:bool = False):
        
        if not comment_use:
            comments = self.doc.get_child_nodes(aw.NodeType.COMMENT, True)
            comments.clear()
        
        self.convert_ext = 'pdf'
        
    # todo: comment 추출 기능 추가
    # todo: 저장 옵션 기능 고려
    # ex)
        # saveOptions = aw.saving.HtmlSaveOptions()
        # saveOptions.export_font_resources = True
        # saveOptions.export_roundtrip_information  = True
            
        # saveOptions = aw.saving.PdfSaveOptions()
        
        # saveOptions = aw.saving.HtmlSaveOptions()
        # saveOptions.allow_embedding_post_script_fonts = True
        
    def compare_two_docx(self, new_file):
        
        newdoc = aw.Document(new_file)
        
        self.doc.compare(newdoc, "test", datetime.today())
        
        print(self.doc.track_revisions)
        print(dir(self.doc))
        
        # for chunk in self.doc.revisions.groups:
        #     print(chunk.text)
        #     print(chunk.revision_type)
        
                
    def convert_document(self, revision:bool = True):
        
        if self.file_ext == '.html':
            builder = aw.DocumentBuilder(self.doc)
            builder.page_setup.paper_size = aw.PaperSize[self.page_setup['paper_size']].value
            builder.page_setup.top_margin = cm_to_point(self.page_setup['top_margin'])
            builder.page_setup.bottom_margin = cm_to_point(self.page_setup['bottom_margin'])
            builder.page_setup.left_margin = cm_to_point(self.page_setup['left_margin'])
            builder.page_setup.right_margin = cm_to_point(self.page_setup['right_margin'])
            builder.page_setup.header_distance = cm_to_point(self.page_setup['header_distance'])
            builder.page_setup.footer_distance = cm_to_point(self.page_setup['footer_distance'])
        
        if revision:
            self.doc.accept_all_revisions()
            
        
            
        save_file_name = f'{self.file_name}.{self.convert_ext}'
            
        if os.path.exists(save_file_name):
            save_file_name = f'{self.file_name}_modi.{self.convert_ext}'           
        
        # saveOptions = aw.saving.HtmlSaveOptions()
        # print(dir(saveOptions))
        # # saveOptions.export_roundtrip_information = True
        # saveOptions.export_document_properties = True
        
        # saveOptions = aw.saving.TxtSaveOptions()
        # self.doc.save(save_file_name, saveOptions)
    
        self.doc.save(save_file_name)
                
                
    def bulk_convert(self, file):
        
        filename, file_extension = os.path.splitext(file)
        # convert to html
        # self.read_document(file)
        # self.doc.accept_all_revisions()
        # self.to_html()
        # self.convert_document()
        
        # # convert to docx
        # self.read_document(filename+".html")
        # self.to_docx()
        # self.convert_document()
        
        # convert to txt
        self.read_document(file)
        self.to_txt()
        self.convert_document()
        

from bs4 import BeautifulSoup as bs

if __name__ == "__main__":
    
    asManager = AsposeManager()
    
    # path ="./testdata/*"
    # file_list = glob.glob(path)
    
    # file_list_docx = [file for file in file_list]
    # for file in file_list_docx:
    #     asManager.bulk_convert(file)    
    
    
    # path ="./계약서/*"
    # folder_list = glob.glob(path)
    
    # for folder in folder_list:
    #     file_list = glob.glob(folder)        
    #     file_list_docx = [file for file in file_list if file.endswith(".DOCX")]
    #     for file in file_list_docx:
    #         asManager.bulk_convert(file)    
            
    # for folder in folder_list:
    #     # folder = folder+'/*'
    #     file_list = glob.glob(folder)        
    #     file_list_docx = [file for file in file_list if file.endswith(".txt")]
    #     for file in file_list_docx:
    #         output = ''
    #         with open(file, encoding='utf-8') as dst_f:
    #             for line in dst_f:

    #                 if not line.isspace():
    #                     output += line
            
    #         filename , ext = os.path.splitext(file)
    #         filename = filename+"_new"
            
    #         with open(filename+ext, 'w', encoding='utf-8') as dst_f:
    #             dst_f.write(output)
    
    # for folder in folder_list:
    #     folder = folder+'/*'
    #     file_list = glob.glob(folder)        
    #     file_list_docx = [file for file in file_list if file.endswith(".docx")]
    #     for file in file_list_docx:
            
    #         os.remove(file)
            # output = ''
            # with open(file, encoding='utf-8') as dst_f:
            #     for line in dst_f:
            #         if not line.isspace():
            #             output += line
            
            # filename , ext = os.path.splitext(file)
            # filename = filename[:-4]
            
            # with open(filename+ext, 'w', encoding='utf-8') as dst_f:
            #     dst_f.write(output)
    
        
    # for file in file_list_docx:
    #     asManager.bulk_convert(file) 
        
    
    
    file = "test홍성우.docx"
    # file = "1.pdf"
    
    doc = aw.Document(file)
    
    doc.accept_all_revisions()
    doc.update_list_labels()
    doc.unlink_fields()
    
    for section in doc:
        section = section.as_section()
        section.headers_footers.clear()
        
        
    
    # html 변환 테스트
    # html_text = doc.to_string(aw.SaveFormat.HTML)
    # html_soup = bs(html_text)
    # html_body = html_soup.find('body')
    
    
    # paragraphs = html_body.find_all('p')
    # for para in paragraphs:
    #     # print(para)
        
    #     if len(para.find_all('a')):
    #         print(para)
        # print(para.get_text())
    
    # comments = html_body.find_all('a')
    # print(comments)
    # print(html_body.prettify())
    # html 변환 테스트

    
    # comments = doc.get_child_nodes(aw.NodeType.COMMENT, True)
    # comments.clear()

    paras = [node.as_paragraph() for node in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)]
    
    # for comment in comments:
    #     print(comment.get_text())
    # print('==================')
    
    # todo: comment, list, line_break, title content 분리
    ret = []
    for para in paras:
        # print(para.to_string(aw.SaveFormat.TEXT))
        
        p = TextParagraph()
        p.text = para.to_string(aw.SaveFormat.TEXT)
                
        comments = para.get_child_nodes(aw.NodeType.COMMENT, True)
         
        if comments.to_array():
            para_comment_ids = [com.as_comment().id for com in comments]
            
            p.comment_ids += para_comment_ids
                    
            comments_range_starts = para.get_child_nodes(aw.NodeType.COMMENT_RANGE_START, True)
            comments_range_ends = para.get_child_nodes(aw.NodeType.COMMENT_RANGE_END, True)
            
            comment_start_id_list = [crs.as_comment_range_start().id for crs in comments_range_starts]
            comment_end_id_list = [cre.as_comment_range_end().id for cre in comments_range_ends]
            
            # print(comment_start_id_list, comment_end_id_list)
            
           
            if not comment_start_id_list:
                                    
                for ce_id in comment_end_id_list:
                    
                    commentStart = doc.get_child(aw.NodeType.COMMENT_RANGE_START, ce_id, True).as_comment_range_start()
                    
                    p_node = commentStart.parent_node.as_paragraph()
                    # print(dir(p_node))
                    
                    while p_node.as_paragraph() != para:
                        # print(p_node.get_text())
                        # print(para.get_text())
                        p_node = p_node.next_sibling
                        # print(p_node.get_text())
                        
                        
            
            if comment_start_id_list:
                
                for cs_id in comment_start_id_list:
                    
                    if cs_id in comment_end_id_list:
                        
                        # print(para.get_text())
                        pass
                    
                    
                    # print(dir(commentStart.parent_node))
                    # print(commentStart.parent_node.node_type_to_string(commentStart.parent_node.node_type))
                    
                    # print(dir(commentStart.next_sibling.as_run()))
                    # print(commentStart.next_sibling.as_run())
                    
        
         
                       
        print(p)
        
        ret.append(
            p
        )
                        
                    
    # print(ret)
    
    # extract_text = doc.to_string(aw.SaveFormat.HTML)
    # extract_text = str(extract_text).replace('<br>', '')
    
    # print(extract_text)
    # doc.save("Output.docx")

    # asManager.read_document(file) 
    
    # # print(asManager.doc.get_text())
    
    # # file = "newtest.docx"
    
    # # asManager.read_document(file) 
    
    # # new_file = "newtest2.html"
    # # asManager.compare_two_docx(new_file)
    
    # asManager.to_docx()
    # asManager.convert_document(revision=False)
    
    
    # convert_documnet
    # revision(검토) 미반영 변환시 False 입력, default는 True(revision 반영)
    # asManager.convert_document(revision=False)

    # html 변환
    # asManager.to_html()
    # asManager.convert_document(revision=False)
    
    # docx 변환
    # asManager.to_docx()
    # asManager.convert_document(revision=False)
    
    # pdf 변환
    # pdf 변환시 comment 포함하여 변환시 True 입력, default는 False(사용안함)
    # asManager.to_pdf(comment_use=False)
    # asManager.convert_document()
