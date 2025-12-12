import zlib
import struct
import zipfile
import xml.etree.ElementTree as ET

class hwpxExtractor(object) :
    def __init__(self, file) :
        self.hwpx = file
        self.namespaces = {
            'opf': 'http://www.idpf.org/2007/opf/',
            'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
            'hp10': 'http://www.hancom.co.kr/hwpml/2016/paragraph',
            'ha' : "http://www.hancom.co.kr/hwpml/2011/app",
            'hs' :"http://www.hancom.co.kr/hwpml/2011/section",
            'hc' : "http://www.hancom.co.kr/hwpml/2011/core" ,
            'hh' : "http://www.hancom.co.kr/hwpml/2011/head" ,
            'hhs' : "http://www.hancom.co.kr/hwpml/2011/history", 
            'hm' : "http://www.hancom.co.kr/hwpml/2011/master-page" ,
            'hpf' : "http://www.hancom.co.kr/schema/2011/hpf" ,
            'dc' : "http://purl.org/dc/elements/1.1/" ,
            'ooxmlchart' : "http://www.hancom.co.kr/hwpml/2016/ooxmlchart" ,
            'hwpunitchar' : "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar",
            'epub' : "http://www.idpf.org/2007/ops" ,
            'config' : "urn:oasis:names:tc:opendocument:xmlns:config:1.0" 
            # 다른 네임스페이스가 있을 수 있습니다. 필요에 따라 추가하세요.
        }

    def extract_text(self):
        try:
            if zipfile.is_zipfile(self.hwpx):
                with zipfile.ZipFile(self.hwpx, 'r') as z:
                    try:
                        with z.open('Contents/section0.xml') as section_xml:
                            tree = ET.parse(section_xml)
                            root = tree.getroot()
                            texts = []
                            for paragraph in root.findall('.//hp:p', self.namespaces):
                                for run in paragraph.findall('.//hp:run', self.namespaces):
                                    element = run.find('hp:t', self.namespaces)
                                    if element is not None and element.text is not None:
                                        texts.append(element.text)
                            return '\n'.join(texts)
                    except KeyError:
                        try:
                            with z.open('BodyText/section0.xml') as section_xml:
                                tree = ET.parse(section_xml)
                                root = tree.getroot()
                                texts = []
                                for paragraph in root.findall('.//hp:p', self.namespaces):
                                    for run in paragraph.findall('.//hp:run', self.namespaces):
                                        element = run.find('hp:t', self.namespaces)
                                        if element is not None and element.text is not None:
                                            texts.append(element.text)
                                return '\n'.join(texts)
                        except KeyError:
                            try:
                                with z.open('Contents/content.hpf') as hpf_file:
                                    tree = ET.parse(section_xml)
                                    root = tree.getroot()
                                    texts = []
                                    for paragraph in root.findall('.//hp:p', self.namespaces):
                                        for run in paragraph.findall('.//hp:run', self.namespaces):
                                            element = run.find('hp:t', self.namespaces)
                                            if element is not None and element.text is not None:
                                                texts.append(element.text)
                                    return '\n'.join(texts)
                            except KeyError:
                                return "ERROR :: .hwpx 텍스트 추출 오류 > section 파일을 찾을 수 없습니다."
            else:
                # ZIP 파일이 아닌 경우 파일을 직접 읽어 처리 (추가 로직 필요)
                return "INFO :: .hwpx 텍스트 추출 > ZIP 파일이 아님. 추가 처리 필요."
        except Exception as e:
            return f"ERROR :: .hwpx 텍스트 추출 오류 > {str(e)}"
            
            