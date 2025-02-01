import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pdf2image import convert_from_path
from PIL import Image, ImageTk
import os
import sys
import time
import threading
import webbrowser

# 중복 실행 방지 플래그
is_running = False

# 타이머 종료 플래그
stop_timer = False  

# 카카오톡 오픈 채팅 링크
kakao_chat_url = "https://open.kakao.com/o/s1o5t4Wg"

# 이미지 클릭 시 오픈채팅 링크 열기
def open_kakao_chat(event):
    webbrowser.open(kakao_chat_url)

def get_unique_folder(base_path):
    """ 같은 폴더명이 있으면 숫자를 붙여 고유한 폴더명 생성 """
    folder_path = base_path
    counter = 1
    while os.path.exists(folder_path):
        folder_path = f"{base_path}_{counter}"
        counter += 1
    return folder_path

# PDF를 PNG로 변환하는 함수
def convert_pdf_to_png(pdf_path, progress_var, progress_label, pdf_button, start_button):
    global is_running, stop_timer
    if is_running:
        return  # 중복 실행 방지

    stop_timer = False # 변환 시작 시 타이머 동작
    is_running = True  # 변환 시작 시 플래그 설정
    pdf_button.config(state=tk.DISABLED) # 변환 시작하면 파일 선택택 버튼 비활성화
    start_button.config(state=tk.DISABLED)  # 변환 시작 버튼 비활성화
    progress_label.config(text="PDF 분석 중...")
    progress_var.set(5)  # 초기 진행률 설정

    try:
        start_time = time.time() # 시작 시간 기록
        update_title(start_time) # 경과시간 업데이트 시작

        # PDF를 이미지로 변환
        images = convert_from_path(pdf_path, dpi=300)
        total_pages = len(images)

        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        pdf_folder = os.path.dirname(pdf_path)

        target_folder = get_unique_folder(os.path.join(pdf_folder, base_name))

        if not os.path.exists(target_folder):
            os.makedirs(target_folder)

        for i, image in enumerate(images):
            progress_label.config(text=f"변환 진행 중... ({i + 1}/{total_pages} 페이지)")
            output_path = os.path.join(target_folder, f"{base_name}_{i + 1}.png")
            image.save(output_path, 'PNG')

            progress = int(((i + 1) / total_pages) * 95) + 5
            progress_var.set(progress)

        end_time = time.time() # 종료 시간 기록
        elapsed_time = end_time - start_time # 소요 시간 계산

        progress_label.config(text="변환 완료!")
        messagebox.showinfo("완료", f"PDF 변환 완료! PNG 파일이 {target_folder}에 저장되었습니다.\n소요 시간: {elapsed_time:.2f}초")
    except Exception as e:
        progress_label.config(text="오류 발생!")
        messagebox.showerror("오류", f"변환 중 오류가 발생했습니다: {str(e)}")
        stop_timer = True # 오류 발생 시 타이머 종료
    finally:
        is_running = False
        stop_timer = True  # 변환 완료 후 타이머 종료
        start_button.config(state=tk.NORMAL)
        pdf_button.config(state=tk.NORMAL)
        progress_var.set(0)  # 프로그레스 바 초기화
        root.title("PDF to PNG 변환기")  # 변환이 끝나면 기본 제목으로 복원

# 변환 진행 중 경과 시간을 업데이트하는 함수
def update_title(start_time):
    def update():
        global stop_timer  # 전역 변수 사용
        while not stop_timer:
            elapsed_time = time.time() - start_time
            root.title(f"PDF to PNG 변환기 - 경과 시간: {elapsed_time:.1f}초")
            time.sleep(1)  # 1초마다 업데이트
        
    threading.Thread(target=update, daemon=True).start()  # 백그라운드 쓰레드 실행

# PDF 파일 선택하는 함수
def select_pdf_file():
    file_path = filedialog.askopenfilename(title="PDF 파일을 선택하세요", filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(0, file_path)
    else:
        messagebox.showwarning("경고", "파일을 선택하지 않았습니다.")

# 변환 시작 (스레드 사용)
def start_conversion():
    global is_running
    if is_running:
        return  # 중복 실행 방지

    pdf_path = pdf_entry.get()
    if not pdf_path:
        messagebox.showwarning("경고", "PDF 파일 경로를 선택해야 합니다.")
        return

    thread = threading.Thread(target=convert_pdf_to_png, args=(pdf_path, progress_var, progress_label, pdf_button, start_button))
    thread.start()

# 실행 파일 내부에서 경로 찾기
def resource_path(relative_path):
    """ PyInstaller로 패키징될 때 실행 파일에서 리소스 경로 찾기 """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return relative_path

# 홍보 이미지 삽입 함수
def add_promo_image():
    image_path = resource_path("abcd.png") # 홍보 이미지 파일 경로
    img = Image.open(image_path)
    img = img.resize((200, 100))  # 적절한 크기로 조절
    img_tk = ImageTk.PhotoImage(img)

    # 이미지 Label 생성
    label = tk.Label(root, image=img_tk, cursor="hand2") # 마우스를 손 모양으로 변경
    label.image = img_tk  # 가비지 컬렉션 방지
    label.grid(pady=10, columnspan=2)

    # 클릭 이벤트 연결
    label.bind("<Button-1>", open_kakao_chat)  # 마우스 왼쪽 클릭 시 open_kakao_chat 함수 실행

# GUI 설정
root = tk.Tk()
root.title("PDF to PNG 변환기")
root.geometry("362x350")


# 레이블
pdf_label = tk.Label(root, text="PDF 파일 선택:")
pdf_label.grid(row=0, column=0, columnspan=2, pady=10)


# 프레임을 사용하여 입력 필드와 버튼을 한 줄에 정렬
file_frame = tk.Frame(root)
file_frame.grid(row=1, column=0, columnspan=2, pady=5, padx=5)

pdf_entry = tk.Entry(file_frame, width=40)
pdf_entry.grid(row=0, column=0, padx=(0, 5))  # 오른쪽 여백 추가

pdf_button = tk.Button(file_frame, text="파일 선택", command=select_pdf_file, background="tomato")
pdf_button.grid(row=0, column=1)

# 진행 상태 레이블
progress_label = tk.Label(root, text="대기 중...", font=("Arial", 10))
progress_label.grid(row=3, column=0,columnspan=2, pady=5)

# 프로그레스 바
progress_var = tk.IntVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=300)
progress_bar.grid(row=4, column=0, columnspan=2, pady=10)

# 변환 시작 버튼
start_button = tk.Button(root, text="변환 시작", command=start_conversion, background="steel blue")
start_button.grid(row=2, column=0, columnspan=2, pady=20)

# 홍보 이미지 추가
add_promo_image()

root.mainloop()