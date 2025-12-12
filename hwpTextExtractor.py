import os
from typing import List, Dict, Any
from io import BytesIO
from PIL import Image
import base64

class HWPTextExtractor:
    """한글 문서에서 텍스트, 이미지, 표를 추출하는 클래스"""
    
    def __init__(self, vlm_endpoint_url: str):
        """
        Args:
            vlm_endpoint_url: VLM API 엔드포인트 URL
        """
        self.vlm_endpoint = vlm_endpoint_url
        self.results = []
    
    def extract_from_hwp(self, hwp_file_path: str) -> List[str]:
        """
        HWP 파일에서 페이지별로 텍스트 추출
        
        Args:
            hwp_file_path: HWP 파일 경로
            
        Returns:
            페이지별 추출된 텍스트 리스트
        """
        try:
            # hwpx 포맷인 경우 (ZIP 기반)
            if hwp_file_path.endswith('.hwpx'):
                return self._extract_from_hwpx(hwp_file_path)
            else:
                # hwp 포맷인 경우 (olefile 기반)
                return self._extract_from_hwp_ole(hwp_file_path)
        except Exception as e:
            print(f"파일 추출 중 오류 발생: {e}")
            return []
    
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
                
                page_results.append("\n".join(page_content))
        
        return page_results
    
    def _extract_from_hwp_ole(self, hwp_path: str) -> List[str]:
        """HWP 파일 처리 (OLE 기반)"""
        import olefile
        
        page_results = []
        
        try:
            ole = olefile.OleFileIO(hwp_path)
            
            # BodyText 스트림에서 섹션 정보 읽기
            sections = [s for s in ole.listdir() if s[0] == 'BodyText']
            
            for section in sections:
                page_content = []
                
                # 텍스트 추출
                try:
                    stream = ole.openstream(section)
                    text = self._parse_hwp_text(stream.read())
                    if text:
                        page_content.append(text)
                except:
                    pass
                
                # 이미지 추출 (BinData에서)
                images = self._extract_images_from_ole(ole)
                for img_data in images:
                    vlm_text = self._process_image_with_vlm(img_data)
                    if vlm_text:
                        page_content.append(vlm_text)
                
                page_results.append("\n".join(page_content))
            
            ole.close()
            
        except Exception as e:
            print(f"HWP 파일 파싱 오류: {e}")
        
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
        # 이미지 참조 찾기
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
                # 표를 이미지로 렌더링 (실제로는 더 복잡한 처리 필요)
                # 여기서는 간단히 표 구조만 추출
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
    
    def _parse_hwp_text(self, data: bytes) -> str:
        """HWP 바이너리에서 텍스트 파싱 (간단한 버전)"""
        try:
            # UTF-16LE로 디코딩 시도
            text = data.decode('utf-16le', errors='ignore')
            # 제어 문자 제거
            text = ''.join(char for char in text if char.isprintable() or char.isspace())
            return text.strip()
        except:
            return ""
    
    def _render_table_to_image(self, table_elem: Any) -> bytes:
        """표를 이미지로 렌더링"""
        # 실제 구현에서는 HTML로 변환 후 screenshot 등을 활용
        # 여기서는 플레이스홀더
        return None
    
    def _process_image_with_vlm(self, img_data: bytes) -> str:
        """VLM을 사용하여 이미지에서 텍스트 추출"""
        import requests
        
        try:
            # 이미지를 base64로 인코딩
            img_base64 = base64.b64encode(img_data).decode('utf-8')
            
            # VLM API 호출
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


# 사용 예시
if __name__ == "__main__":
    # VLM 엔드포인트 URL 설정
    VLM_ENDPOINT = "https://your-vlm-api-endpoint.com/extract"
    
    # 추출기 초기화
    extractor = HWPTextExtractor(VLM_ENDPOINT)
    
    # HWP 파일에서 텍스트 추출
    hwp_file = "sample.hwp"
    page_results = extractor.extract_from_hwp(hwp_file)
    
    # 결과 출력
    for i, page_content in enumerate(page_results):
        print(f"\n=== 페이지 {i+1} ===")
        print(page_content)
        print(f"\narr[{i}] = {repr(page_content[:100])}...")  # 처음 100자만 표시