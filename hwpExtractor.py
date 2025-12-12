import re
import zlib
import struct
import olefile
import zipfile
import unicodedata

class hwpExtractor(object) :
    def __init__(self, file) :
        self.file = file

    def remove_chinese_characters(self, s: str):
        return re.sub(r'[\u4e00-\u9fff]+', '', s)

    def remove_control_characters(self, s):
        return "".join(ch for ch in s if unicodedata.category(ch)[0] != "C")

    def clean_text(self, s: str):
        s = self.remove_chinese_characters(s)
        s = self.remove_control_characters(s)
        return s
    
    def get_text(self) :
        try :
            ole_file = olefile.OleFileIO(self.file)
        except Exception as e:
            return self.get_hwp5_text()

        dirs = ole_file.listdir()

        if ["FileHeader"] not in dirs or ["\x05HwpSummaryInformation"] not in dirs:
            raise Exception(f"ERROR :: HWPEXTRACTOR :: {self.file.name} is Not a valid HWP OLE file")
        
        header = ole_file.openstream("FileHeader")
        header_data = header.read()
        is_zipped = (header_data[36] & 1) == 1

        section_nums = []
        for d in dirs:
            if d[0] == "BodyText":
                section_nums.append(int(d[1][len("Section"):]))
        sections = ["BodyText/Section" + str(x) for x in sorted(section_nums)]

        text = ""

        for section in sections:
            bodytext = ole_file.openstream(section)
            data = bodytext.read()
            if is_zipped:
                try:
                    unzipped_data = zlib.decompress(data, -15)
                except zlib.error as e:
                    print(f"ERROR :: HWPEXTRACTOR :: {self.file.name} 압축 해제 실패: {str(e)} >> 일부만 압축 해제 시도 ... ")
                    unzipped_data = self.partial_decompress(data)
            else:
                unzipped_data = data

            section_text = ""
            i = 0
            size = len(unzipped_data)

            while i < size :
                header = struct.unpack_from("<I", unzipped_data, i)[0]
                rec_type = header & 0x3ff
                rec_len = (header >> 20) & 0xfff

                if rec_type in [51] :
                    i += 4 + rec_len
                    continue

                if rec_type in [67] :
                    rec_data = unzipped_data[i+4:i+4+rec_len]
                    section_text += self.decode_text(rec_data)
                    section_text += "\n"

                i += 4 + rec_len

            cleaned_text = self.clean_text(section_text)
            text += cleaned_text
            text += "\n"

            return text
        
    def decode_text(self, rec_data):
        for encoding in ['utf-16', 'utf-8', 'cp949', 'euc-kr']:
            try:
                return rec_data.decode(encoding)
            except UnicodeDecodeError as e:
                continue
        
    def get_hwp5_text(self) :
        if not zipfile.is_zipfile(self.file):
            raise Exception(f"ERROR :: HWP5EXTRACTOR :: {self.file.name}는 hwp 5.0 파일이 아닙니다.")
        
        with zipfile.ZipFile(self.file, 'r') as hwp_zip:
            sections = [f for f in hwp_zip.namelist() if f.startswith("Contents/section")]

            if not sections :
                raise Exception(f"ERROR :: HWP5EXTRACTOR :: {self.file.name} >> hwp5를 시도했으나 section을 찾을 수 없습니다.")
            
            text = ""
            for section in sections:
                with hwp_zip.open(section) as bodytext:
                    data = bodytext.read()
                    try:
                        unzipped_data = zlib.decompress(data, -15)
                    except zlib.error:
                        unzipped_data = data

                    section_text = ""
                    i = 0
                    size = len(unzipped_data)
                    while i < size:
                        header = struct.unpack_from("<I", unzipped_data, i)[0]
                        rec_type = header & 0x3ff
                        rec_len = (header >> 20) & 0xfff

                        if rec_type in [51] :
                            i += 4 + rec_len
                            continue

                        if rec_type in [67] :
                            rec_data = unzipped_data[i+4:i+4+rec_len]
                            section_text += self.decode_text(rec_data)
                            section_text += "\n"

                        i += 4 + rec_len

                    cleaned_text = self.clean_text(section_text)
                    text += cleaned_text
                    text += "\n"

            return text
        