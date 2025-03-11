import os
import re
import win32com.client as win32

def split_docx(input_path, output_dir, word_app, q_a="q"):

    # === 사용자 환경 설정 ===
    WORK_DIR = os.getcwd()                                     # 현재 작업 디렉터리
    # WORD_FILES_DIR = os.path.join(WORK_DIR, "word_files")      # 분할할 .docx 파일(문제/정답)이 있는 폴더
    WORD_FILES_DIR = os.path.join(WORK_DIR, "word_files/test_word")      # 분할할 .docx 파일(문제/정답)이 있는 폴더
    OUTPUT_DIR     = os.path.join(WORK_DIR, "split")         # 분할된 파일을 저장할 폴더

    # 문제/정답 패턴 정의
    QUESTION_START_PATTERN = re.compile(r'^[0-9]+\.')
    QUESTION_END_PATTERN   = re.compile(r'^\[1\]$')

    """
    하나의 함수로 문제/정답 분할을 모두 처리.
    
    :param input_path:  분할할 Word 파일의 전체 경로
    :param output_dir:  분할된 결과물을 저장할 폴더
    :param word_app:    Word Application COM 객체 (pywin32)
    :param q_a:         "q"면 문제, "a"면 정답 모드
                       - "q" => '시작 패턴' + '[1]'(문제 끝) 을 기준으로 분할
                       - "a" => '시작 패턴'만 보고, 다음 '시작 패턴' 직전까지를 한 블록
    """
    # Word 문서 열기
    doc = word_app.Documents.Open(input_path)

    # 분할이 끝난 파일들의 이름 list
    result_file_list = []
    
    try:
        paragraphs = doc.Paragraphs
        
        # 결과 저장을 위한 (start_par_idx, end_par_idx) 목록
        blocks = []
        current_start = None
        
        # 문단 순회
        for i in range(1, paragraphs.Count + 1):
            p_text = paragraphs(i).Range.Text.strip()
            
            # 문제/정답 "시작" 패턴 매칭
            if QUESTION_START_PATTERN.match(p_text):
                # 만약 이미 시작된 블록이 있었다면, 지금 문단(i) 직전까지를 끝으로 확정
                if q_a == "a" and current_start is not None:
                    # "정답" 모드는 [1]이 아니라, "다음 문제 번호"가 나오면 이전 문제를 끝냄
                    blocks.append((current_start, i - 1))
                    current_start = None
                
                # 새 블록 시작 인덱스 기록
                current_start = i
            
            # 만약 "문제 모드(q)"이고, 문제 끝 패턴([1]) 매칭되면
            if q_a == "q":
                if QUESTION_END_PATTERN.match(p_text):
                    # 현재 블록이 시작된 상태라면, 여기까지를 블록으로 확정
                    if current_start is not None:
                        blocks.append((current_start, i))
                        current_start = None
        
        # 마지막 블록 처리
        if current_start is not None:
            # "문제(q)" 모드는 [1]로 끝난 블록이 아닐 수도 있으므로, 문서 끝까지 포함
            # "정답(a)" 모드도 마찬가지로, 다음 문제 시작이 없었다면 끝까지 포함
            blocks.append((current_start, paragraphs.Count))
        
        # 분할할 블록이 없다면 종료
        if not blocks:
            print(f"[{os.path.basename(input_path)}] => 분할할 블록(패턴)이 없습니다.")
            return
        
        # 결과물 저장 폴더가 없으면 생성
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 원본 파일명(확장자 제외)
        base_filename = os.path.splitext(os.path.basename(input_path))[0].split("_")
    
        block_count = 0
        
        # 각 블록마다 새 문서를 생성해 붙여넣기
        for start_par_idx, end_par_idx in blocks:
            start_range = paragraphs(start_par_idx).Range.Start
            end_range   = paragraphs(end_par_idx).Range.End
            copy_range  = doc.Range(Start=start_range, End=end_range)
            
            new_doc = word_app.Documents.Add()
            copy_range.Copy()
            new_doc.Range().Paste()
            
            # [선택] 마지막 빈 단락 제거
            while new_doc.Paragraphs.Count > 0:
                last_par_text = new_doc.Paragraphs(new_doc.Paragraphs.Count).Range.Text.strip()
                if last_par_text == "":
                    new_doc.Paragraphs(new_doc.Paragraphs.Count).Range.Delete()
                else:
                    break
            
            # 첫 줄 확인 (디버그용)
            if new_doc.Paragraphs.Count > 0:
                first_line_text = new_doc.Paragraphs(1).Range.Text.strip()
            else:
                first_line_text = ""
            
            block_count += 1

            # "원본파일명_01.docx", "원본파일명_02.docx" 형태로 저장
            if q_a == "q":
                new_filename = f"{'_'.join(base_filename[:-1])}_P{block_count:02d}.docx"
                
            elif q_a == "a":
                new_filename = f"{'_'.join(base_filename[:-2])}_P{block_count:02d}_MS.docx"

            ### result_file_list에 분할된 문제의 filename을 append
            result_file_list.append(new_filename)

            
            save_path = os.path.join(output_dir, new_filename)
            print(f"save path : {save_path}")
            new_doc.SaveAs2(save_path, FileFormat=16)  # 16 = wdFormatXMLDocument(.docx)
            new_doc.Close()
            
            print(f"- {block_count:02d}번째 분할: '{new_filename}' (첫 줄: {first_line_text})")
        
        print(f"==> [{os.path.basename(input_path)}] 총 {block_count}개로 분할 완료. (모드='{q_a}')\n")
    
    finally:
        ### 결과 확인
        print(f"분할된 문제 list : {result_file_list}")
        doc.Close(False)
        return result_file_list
