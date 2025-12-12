import os
import struct
from typing import List, Dict, Any, Tuple
from io import BytesIO
from PIL import Image
import base64

class HWPTextExtractor:
    """한글 문서에서 텍스트, 이미지, 표를 추출하는 클래스"""
    
    # HWP 버전 상수
    HWP_VERSION_UNKNOWN = 0
    HWP_VERSION_3X = 3
    HWP_VERSION_5X = 5
    HWP_VERSION_HWPX = 6
    
    def __init__(self, vlm_endpoint_url: str):
        """
        Args:
            vlm_endpoint_url: VLM API 엔드포인트 URL
        """
        self.vlm_endpoint = vlm_endpoint_url
        self.results = []
        self.file_version = self.HWP_VERSION_UNKNOWN
    
    def detect_hwp_version(self, file_path: str) -> int:
        """
        HWP 파일의 버전을 감지
        
        Args:
            file_path: HWP 파일 경로
            
        Returns:
            버전 코드 (3: HWP 3.x, 5: HWP 5.x, 6: HWPX)
        """
        # 확장자로 먼저 확인
        if file_path.lower().endswith('.hwpx'):
            return self.HWP_VERSION_HWPX
        
        try:
            with open(file_path, 'rb') as f:
                # 파일 시그니처 읽기 (처음 32바이트)
                signature = f.read(32)
                
                # HWP 5.0 이상 시그니처 확인 (CFB/OLE 파일)
                # CFB 파일은 0xD0CF11E0A1B11AE1로 시작
                if signature[:8] == b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1':
                    # OLE 파일이므로 HWP 5.0 확인
                    f.seek(0)
                    return self._verify_hwp5_signature(f)
                
                # HWP 3.0 시그니처 확인
                # HWP 3.0은 "HWP Document File"로 시작
                if signature[:16] == b'HWP Document Fil':
                    return self.HWP_VERSION_3X
                
                # 기타 확인
                return self.HWP_VERSION_UNKNOWN
                
        except Exception as e:
            print(f"버전 감지 중 오류 발생: {e}")
            return self.HWP_VERSION_UNKNOWN
    
    def _verify_hwp5_signature(self, f) -> int:
        """
        HWP 5.0 파일의 시그니처를 상세히 확인
        
        Args:
            f: 파일 객체
            
        Returns:
            버전 코드
        """
        try:
            import olefile
            
            f.seek(0)
            ole = olefile.OleFileIO(f)
            
            # FileHeader 스트림 확인
            if ole.exists('FileHeader'):
                header_stream = ole.openstream('FileHeader')
                header_data = header_stream.read()
                
                # HWP 5.0 시그니처 확인 (32바이트)
                # 서명: "HWP Document File" (앞 16바이트)
                if header_data[:16] == b'HWP Document Fil':
                    # 버전 정보 읽기 (offset 16~20)
                    version_bytes = header_data[16:20]
                    if len(version_bytes) == 4:
                        version = struct.unpack('<I', version_bytes)[0]
                        major = (version >> 24) & 0xFF
                        minor = (version >> 16) & 0xFF
                        micro = (version >> 8) & 0xFF
                        build = version & 0xFF
                        
                        print(f"HWP 버전 감지: {major}.{minor}.{micro}.{build}")
                        
                        if major == 5:
                            return self.HWP_VERSION_5X
                
            ole.close()
            return self.HWP_VERSION_5X
            
        except Exception as e:
            print(f"HWP 5.0 시그니처 확인 실패: {e}")
            return self.HWP_VERSION_5X
    
    def extract_from_hwp(self, hwp_file_path: str) -> List[str]:
        """
        HWP 파일에서 페이지별로 텍스트 추출
        
        Args:
            hwp_file_path: HWP 파일 경로
            
        Returns:
            페이지별 추출된 텍스트 리스트
        """
        # 버전 감지
        self.file_version = self.detect_hwp_version(hwp_file_path)
        
        print(f"감지된 파일 버전: {self._get_version_name()}")
        
        try:
            # 버전에 따른 처리
            if self.file_version == self.HWP_VERSION_HWPX:
                return self._extract_from_hwpx(hwp_file_path)
            elif self.file_version == self.HWP_VERSION_5X:
                return self._extract_from_hwp5(hwp_file_path)
            elif self.file_version == self.HWP_VERSION_3X:
                return self._extract_from_hwp3(hwp_file_path)
            else:
                print("알 수 없는 HWP 파일 형식입니다.")
                return []
                
        except Exception as e:
            print(f"파일 추출 중 오류 발생: {e}")
            return []
    
    def _get_version_name(self) -> str:
        """버전 코드를 문자열로 변환"""
        version_names = {
            self.HWP_VERSION_UNKNOWN: "Unknown",
            self.HWP_VERSION_3X: "HWP 3.0 (구버전)",
            self.HWP_VERSION_5X: "HWP 5.0 (OLE/CFB)",
            self.HWP_VERSION_HWPX: "HWPX (XML 기반)"
        }
        return version_names.get(self.file_version, "Unknown")
    
    def _extract_from_hwp3(self, hwp3_path: str) -> List[str]:
        """
        HWP 3.0 파일 처리
        HWP 3.0은 단순한 바이너리 구조
        """
        page_results = []
        
        try:
            with open(hwp3_path, 'rb') as f:
                # 헤더 스킵 (256바이트)
                header = f.read(256)
                
                # 문서 정보 구조체 파싱
                # HWP 3.0은 고정된 구조를 가지고 있음
                doc_info = f.read(256)
                
                # 본문 데이터 읽기
                content = f.read()
                
                # 텍스트 추출 시도
                # HWP 3.0은 주로 EUC-KR 또는 CP949 인코딩 사용
                try:
                    text = content.decode('cp949', errors='ignore')
                    # 제어 문자 제거
                    text = ''.join(char for char in text if char.isprintable() or char.isspace())
                    
                    # 간단히 한 페이지로 처리 (3.0은 페이지 구분이 명확하지 않음)
                    if text.strip():
                        page_results.append(text.strip())
                except Exception as e:
                    print(f"HWP 3.0 텍스트 디코딩 실패: {e}")
                
        except Exception as e:
            print(f"HWP 3.0 파일 파싱 오류: {e}")
        
        return page_results if page_results else [""]
    
    def _extract_from_hwp5(self, hwp5_path: str) -> List[str]:
        """
        HWP 5.0 파일 처리 (OLE 기반)
        """
        import olefile
        import zlib
        
        page_results = []
        
        try:
            ole = olefile.OleFileIO(hwp5_path)
            
            # FileHeader 읽기
            if ole.exists('FileHeader'):
                header_stream = ole.openstream('FileHeader')
                header_data = header_stream.read()
                
                # 압축 여부 확인 (속성 플래그)
                flags = struct.unpack('<I', header_data[36:40])[0]
                is_compressed = bool(flags & 0x01)
                
                print(f"문서 압축 여부: {'압축됨' if is_compressed else '압축 안 됨'}")
            
            # BodyText 섹션 처리
            sections = [s for s in ole.listdir() if s[0] == 'BodyText']
            
            for section in sorted(sections):
                page_content = []
                
                try:
                    stream = ole.openstream(section)
                    data = stream.read()
                    
                    # 압축된 경우 압축 해제
                    if is_compressed:
                        try:
                            data = zlib.decompress(data, -15)
                        except:
                            pass
                    
                    # 텍스트 추출
                    text = self._parse_hwp5_text(data)
                    if text:
                        page_content.append(text)
                    
                except Exception as e:
                    print(f"섹션 {section} 처리 중 오류: {e}")
                
                # 이미지 추출 및 VLM 처리
                images = self._extract_images_from_ole(ole)
                for img_data in images:
                    vlm_text = self._process_image_with_vlm(img_data)
                    if vlm_text:
                        page_content.append(vlm_text)
                
                page_results.append("\n".join(page_content) if page_content else "")
            
            ole.close()
            
        except Exception as e:
            print(f"HWP 5.0 파일 파싱 오류: {e}")
        
        return page_results if page_results else [""]
    
    def _parse_hwp5_text(self, data: bytes) -> str:
        """
        HWP 5.0 바이너리에서 텍스트 파싱
        HWP 5.0은 레코드 구조로 되어 있음
        """
        texts = []
        offset = 0
        
        try:
            while offset < len(data) - 4:
                # 레코드 헤더 읽기 (4바이트)
                if offset + 4 > len(data):
                    break
                
                record_header = struct.unpack('<I', data[offset:offset+4])[0]
                offset += 4
                
                # 레코드 태그 및 레벨, 크기 추출
                tag_id = record_header & 0x3FF
                level = (record_header >> 10) & 0x3FF
                size = (record_header >> 20) & 0xFFF
                
                if size == 0xFFF:  # 확장 크기
                    if offset + 4 > len(data):
                        break
                    size = struct.unpack('<I', data[offset:offset+4])[0]
                    offset += 4
                
                # 레코드 데이터 읽기
                if offset + size > len(data):
                    break
                
                record_data = data[offset:offset+size]
                offset += size
                
                # 텍스트 레코드 처리 (HWPTAG_PARA_TEXT = 67)
                if tag_id == 67:
                    try:
                        # UTF-16LE로 디코딩
                        text = record_data.decode('utf-16le', errors='ignore')
                        # 제어 문자 제거
                        text = ''.join(char for char in text if char.isprintable() or char.isspace())
                        if text.strip():
                            texts.append(text.strip())
                    except:
                        pass
                        
        except Exception as e:
            print(f"레코드 파싱 중 오류: {e}")
        
        return ' '.join(texts)
    
    def _extract_from_hwpx(self, hwpx_path: str) -> List[str]:
        """HWPX 파일 처리 (ZIP 기반)"""
        import zipfile
        from xml.etree import ElementTree as ET
        
        page_results = []
        
        with zipfile.ZipFile(hwpx_path, 'r') as hwpx:
            # Contents 디렉토리에서 섹션별 XML 파싱
            section_files = [f for f in hwpx.namelist() if f.startswith('Contents/section')]
            
            for section_file in sorted(section_files):
                page_content = []
                
                xml_content = hwpx.read(section_file)
                root = ET.fromstring(xml_content)
                
                # 텍스트 추출
                text_content = self._extract_text_from_xml(root)
                if text_content:
                    page_content.append(text_content)
                
                # 이미지 추출 및 VLM 처리
                images = self._extract_images_from_section(hwpx, root)
                for img_data in images:
                    vlm_text = self._process_image_with_vlm(img_data)
                    if vlm_text:
                        page_content.append(vlm_text)
                
                # 표 추출 및 VLM 처리
                tables = self._extract_tables_from_xml(hwpx, root)
                for table_img in tables:
                    vlm_text = self._process_image_with_vlm(table_img)
                    if vlm_text:
                        page_content.append(vlm_text)
                
                page_results.append("\n".join(page_content) if page_content else "")
        
        return page_results
    
    def _extract_text_from_xml(self, root: Any) -> str:
        """XML에서 텍스트 추출"""
        texts = []
        # hp:t 태그에서 텍스트 추출
        for text_elem in root.iter():
            if text_elem.tag.endswith('t') and text_elem.text:
                texts.append(text_elem.text)
        return " ".join(texts)
    
    def _extract_images_from_section(self, hwpx_zip, root: Any) -> List[bytes]:
        """섹션에서 이미지 추출"""
        images = []
        for img_elem in root.iter():
            if 'Pic' in img_elem.tag or 'Image' in img_elem.tag:
                img_id = img_elem.get('BinItemID') or img_elem.get('href')
                if img_id:
                    try:
                        img_path = f'bindata/{img_id}'
                        if img_path in hwpx_zip.namelist():
                            img_data = hwpx_zip.read(img_path)
                            images.append(img_data)
                    except:
                        pass
        return images
    
    def _extract_tables_from_xml(self, hwpx_zip, root: Any) -> List[bytes]:
        """표를 이미지로 변환하여 추출"""
        table_images = []
        
        for table_elem in root.iter():
            if 'tbl' in table_elem.tag.lower():
                table_img = self._render_table_to_image(table_elem)
                if table_img:
                    table_images.append(table_img)
        
        return table_images
    
    def _extract_images_from_ole(self, ole) -> List[bytes]:
        """OLE 파일에서 이미지 추출"""
        images = []
        
        try:
            bin_data_dirs = [d for d in ole.listdir() if d[0] == 'BinData']
            for bin_dir in bin_data_dirs:
                try:
                    stream = ole.openstream(bin_dir)
                    img_data = stream.read()
                    images.append(img_data)
                except:
                    pass
        except:
            pass
        
        return images
    
    def _render_table_to_image(self, table_elem: Any) -> bytes:
        """표를 이미지로 렌더링 (플레이스홀더)"""
        # 실제 구현에서는 HTML로 변환 후 screenshot 등을 활용
        return None
    
    def _process_image_with_vlm(self, img_data: bytes) -> str:
        """VLM을 사용하여 이미지에서 텍스트 추출"""
        import requests
        
        try:
            img_base64 = base64.b64encode(img_data).decode('utf-8')
            
            response = requests.post(
                self.vlm_endpoint,
                json={
                    "image": img_base64,
                    "task": "text_extraction"
                },
                timeout=30
            )
            
            if response.status_code == 200:
                result = response.json()
                return result.get('text', '')
            else:
                print(f"VLM API 오류: {response.status_code}")
                return ""
                
        except Exception as e:
            print(f"VLM 처리 중 오류: {e}")
            return ""
    
    def get_page_content(self, page_index: int) -> str:
        """특정 페이지의 내용 반환"""
        if 0 <= page_index < len(self.results):
            return self.results[page_index]
        return ""
    
    def get_file_info(self) -> Dict[str, Any]:
        """파일 정보 반환"""
        return {
            "version": self._get_version_name(),
            "version_code": self.file_version,
            "total_pages": len(self.results)
        }


# 사용 예시
if __name__ == "__main__":
    # VLM 엔드포인트 URL 설정
    VLM_ENDPOINT = "https://your-vlm-api-endpoint.com/extract"
    
    # 추출기 초기화
    extractor = HWPTextExtractor(VLM_ENDPOINT)
    
    # HWP 파일에서 텍스트 추출
    hwp_file = "sample.hwp"
    page_results = extractor.extract_from_hwp(hwp_file)
    
    # 파일 정보 출력
    info = extractor.get_file_info()
    print(f"\n파일 정보:")
    print(f"- 버전: {info['version']}")
    print(f"- 총 페이지: {info['total_pages']}")
    
    # 결과 출력
    for i, page_content in enumerate(page_results):
        print(f"\n=== 페이지 {i+1} ===")
        print(page_content[:200] if len(page_content) > 200 else page_content)
        print(f"\narr[{i}] 길이: {len(page_content)} 문자")