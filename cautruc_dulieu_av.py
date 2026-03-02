class ExamQuestionAV:
    """Đại diện cho 1 câu hỏi con trong môn AV/ĐGNL"""
    def __init__(self):
        self.label = ""       # VD: "Question 1", "Câu 1"
        self.number = 0       # Số thứ tự (int)
        self.stem = []        # Nội dung dẫn của câu (List paragraphs)
        self.options = {'A': [], 'B': [], 'C': [], 'D': []} # Các phương án
        self.solution = []    # Lời giải
        self.key_value = ""   # Đáp án (A, B... hoặc text tự luận)
        self.type = "MCQ"     # MCQ (Trắc nghiệm) hoặc ESSAY (Tự luận)

class ExamCluster:
    """Đại diện cho 1 nhóm câu hỏi (Bài đọc, Biểu đồ...)"""
    def __init__(self):
        self.start_marker = None # Paragraph chứa #*#
        self.end_marker = None   # Paragraph chứa #**#
        self.context = []        # Nội dung đoạn văn/biểu đồ chung (List paragraphs)
        self.questions = []      # Danh sách các ExamQuestionAV thuộc nhóm này
        self.group_type = "NORMAL" # NORMAL (Thường) hoặc FIXED (Cố định - #@#)