import os

class Config:
    # 默认文件路径
    DEFAULT_QUESTION_BANK = "题库.xlsx"
    DEFAULT_RESULTS_FILE = "考核结果.xlsx"
    
    # 评分标准
    SCORING_CATEGORIES = [
        "知识点分数",
        "回答流利程度", 
        "对知识点的扩展",
        "回答思路是否严谨"
    ]
    
    MAX_SCORE_PER_CATEGORY = 25
    MAX_TOTAL_SCORE = 100
    
    # 界面设置
    WINDOW_TITLE = "题库随机抽题评分系统"
    WINDOW_SIZE = (1200, 800)
    
    @staticmethod
    def get_data_dir():
        """获取数据目录"""
        data_dir = os.path.join(os.path.expanduser("~"), "题库评分系统")
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)
        return data_dir