import pandas as pd
import random

# 데이터 프레임 생성
data = {
    '이름': ['학생{}'.format(i) for i in range(1, 26)],
    '비밀번호': ['password{}'.format(i) for i in range(1, 26)],
    '독서': [random.randint(0, 100) for _ in range(25)],
    '수학II': [random.randint(0, 100) for _ in range(25)],
    '영어II': [random.randint(0, 100) for _ in range(25)],
    '물리I': [random.randint(0, 100) for _ in range(25)],
    '화학I': [random.randint(0, 100) for _ in range(25)],
    '생명I': [random.randint(0, 100) for _ in range(25)],
    '한국지리': [random.randint(0, 100) for _ in range(25)],
    '동아시아사': [random.randint(0, 100) for _ in range(25)],
    '경제': [random.randint(0, 100) for _ in range(25)],
    '정치와 법': [random.randint(0, 100) for _ in range(25)],
    '일본어 I': [random.randint(0, 100) for _ in range(25)],
}

# 데이터프레임 생성
df = pd.DataFrame(data)

# 엑셀 파일로 저장
df.to_excel('성적표.xlsx', index=False)
