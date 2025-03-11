import os
import pandas as pd


def add_categories_to_csv(category_csv_path, file_list, exe_dir):
    """  
    file_list : csv에 추가할 file들의 이름들을 포함하는 list.
    """
    # 문제 형식[1A, 1B, 2]
    # 1A : 객관식 
    # 1B : data 해석 문제 
    # 2 : 주관식
    question_type_set = {"1A", "1B", "2", "A1", "B1"}

    # SL -> 레벨 나눠놓은거 
    # SL : standard level 탐구로 치면 사탐1 과탐1 정도 
    # HL : high level 사탐2 과탐2 정도 느낌
    question_level_set = {"SL", "HL"}

    # TZ0 -> Timezone 
    # time zone 이라고 전세계가 시험보는거라 시차에 따라서 
    time_zone_set = {"TZ0", "TZ1", "TZ2"}

    # 과목
    # 맨 처음([0])
    subject_set = {"Phy", "Math"}

    # 문제 번호
    # 맨 마지막([-1])
    # MS일 경우 답지

    if os.path.exists(category_csv_path):
        print("csv파일 존재함.")
        df = pd.read_csv(category_csv_path)
    else:
        print("csv파일 없음.")
        df = pd.DataFrame(columns=["file_name", "subject", "q_number", "q_type", "q_level", "time_zone"])

    df_file_list = set(df["file_name"])

    count = 0

    for file in file_list:
        if file in df_file_list:
            continue

        full_file_name = file

        split_file_name_os = os.path.splitext(file)[0].split("_")

        ### 정답파일은 넘어가기
        if split_file_name_os[-1] == "MS":
            continue

        # 기본값을 설정하여 변수 미할당 오류 방지
        question_type = None
        question_level = None
        time_zone = None
        
        # 과목
        subject = split_file_name_os[0]
        # 문제 번호호
        question_number = split_file_name_os[-1]

        rows = []

        for idx in range(1, len(split_file_name_os)-1):
            # 문제 타입
            if split_file_name_os[idx] in question_type_set:
                question_type = split_file_name_os[idx]
            # 문제 레벨
            elif split_file_name_os[idx] in question_level_set:
                question_level = split_file_name_os[idx]
            # 타임존
            elif split_file_name_os[idx] in time_zone_set:
                time_zone = split_file_name_os[idx]


        rows.append({
            "file_name":full_file_name, 
            "subject":subject, 
            "q_number":question_number, 
            "q_type":question_type, 
            "q_level":question_level,
            "time_zone":time_zone, 
            # "a_file_name": "".join(["_".join(split_file_name_os), "_MS.docx"])
        })

        df = pd.concat([df, pd.DataFrame(rows)], ignore_index=True)
        count += 1

    ### 임시
    folder_path = os.path.join(exe_dir, "csv_file")  # 상대경로 (현재 작업 디렉토리 기준)

    if not os.path.exists(folder_path):  # 폴더 존재 여부 확인
        os.makedirs(folder_path)  # 폴더 생성

    df.to_csv(category_csv_path, index=False) # .exe 경로
    # df.to_csv(r"D:\YearDreamSchool-D\python_projects\msword_pjt\csv_file\categories.csv", index=False) # .py 경로

    print(f"{count}개의 파일이 추가되었습니다.")