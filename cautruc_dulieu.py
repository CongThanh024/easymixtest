import re

class Pattern:
    # 1. BẮT ĐẦU CÂU: "Câu 1.", "Câu 1:"
    # Bắt buộc nằm đầu dòng hoặc sau khoảng trắng
    START_CAU = r"(?:^|\s|#)(@|!)?(?:Câu|Question)\s+(\d+)(?:[\.:])"

    # 2. PHẦN THI
    START_PHAN = r"(?:^|\n)\s*(PHẦN\s+[IVX]+)(?:[\.:])"

    # 3. PHƯƠNG ÁN TN: "A.", "B."...
    OPT_TN = r"(?:^|\s)([A-D])\."

    # 4. Ý ĐÚNG SAI: "a)", "b)"...
    OPT_DS = r"(?:^|\s)([a-d])\)"

    # 5. TAG KEY: <key=...>
    TAG_KEY = r"(?:^|\s)(<key=.*?>)"

    # 6. TAG TỰ LUẬN
    TAG_TL = r"(?:^|\s)(<(?:Tự luận|TU LUAN|tu luan)>)"
    
    # 7. LỜI GIẢI: "Lời giải", "Hướng dẫn giải" (Không bắt buộc dấu chấm/hai chấm ngay sau, cứ thấy là bắt)
    # Thường nó đứng đầu dòng
    LOI_GIAI = r"^(?:Lời giải|LỜI GIẢI|HƯỚNG DẪN GIẢI|Hướng dẫn giải|HD|LG)(?:[\.:])?"

    # 8. RÁC CHỌN ĐÁP ÁN: "Chọn A", "Chọn B."...
    # Bắt buộc đứng đầu câu hoặc sau dấu chấm/phẩy
    RAC_CHON = r"(?:^|[\.,\s])(Chọn\s+[A-D])(?:[\.:]?)"

    # 9. NHÓM
    GROUP_OPEN_THUONG = r"#\*#"
    GROUP_OPEN_CODINH = r"#@#"
    GROUP_CLOSE = r"#\*\*#"

class CauHoi:
    def __init__(self):
        self.id_goc = 0
        self.noi_dung = [] 
        self.ds_phuong_an = {'A': [], 'B': [], 'C': [], 'D': []}
        self.ds_y_dung_sai = {'a': [], 'b': [], 'c': [], 'd': []}
        self.loi_giai = []
        self.loai_cau = "TN"
        self.thuoc_nhom = ""
        self.ghim_cho = False
        self.ghim_phuong_an = False
        self.dap_an_dung = []
        self.co_key = False

    def them_noi_dung(self, para_obj):
        self.noi_dung.append(para_obj)