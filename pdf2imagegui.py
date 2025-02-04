import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pdf2image import convert_from_path
from PIL import Image, ImageTk
import os
import sys
import time
import threading
import multiprocessing
import webbrowser
import gc

is_running = False
stop_timer = False  

kakao_chat_url = "https://open.kakao.com/o/s1o5t4Wg"

def open_kakao_chat(event):
    webbrowser.open(kakao_chat_url)

def get_unique_folder(base_path):
    folder_path = base_path
    counter = 1
    while os.path.exists(folder_path):
        folder_path = f"{base_path}_{counter}"
        counter += 1
    return folder_path

def convert_pages(args):
    pdf_path, start, end, target_folder, base_name = args
    images = convert_from_path(pdf_path, dpi=300, first_page=start, last_page=end, thread_count=1)
    
    for i, image in enumerate(images, start=start):
        output_path = os.path.join(target_folder, f"{base_name}_{i}.png")
        image.save(output_path, 'PNG')
        del image
        gc.collect()

def convert_pdf_to_png(pdf_path, progress_var, progress_label, pdf_button, start_button):
    global is_running, stop_timer
    if is_running:
        return

    stop_timer = False
    is_running = True
    pdf_button.config(state=tk.DISABLED)
    start_button.config(state=tk.DISABLED)
    progress_label.config(text="PDF 분석 중...")
    progress_var.set(5)

    try:
        start_time = time.time()
        update_title(start_time)

        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        pdf_folder = os.path.dirname(pdf_path)
        target_folder = get_unique_folder(os.path.join(pdf_folder, base_name))
        os.makedirs(target_folder, exist_ok=True)

        total_pages = len(convert_from_path(pdf_path, dpi=50)) # 페이지 변환할 땐 고해상도가 필요없어서 50으로함함
        
        num_processes = max(1, multiprocessing.cpu_count() // 2)
        page_chunk = 2
        page_ranges = [(pdf_path, i + 1, min(i + page_chunk, total_pages), target_folder, base_name) 
                       for i in range(0, total_pages, page_chunk)]
        
        with multiprocessing.Pool(processes=num_processes) as pool:
            pool.map(convert_pages, page_ranges)

        end_time = time.time()
        elapsed_time = end_time - start_time

        progress_label.config(text="변환 완료!")
        print(f"소요시간 : {elapsed_time}")
        messagebox.showinfo("완료", f"PDF 변환 완료! PNG 파일이 {target_folder}에 저장되었습니다.\n소요 시간: {elapsed_time:.2f}초")
    except Exception as e:
        progress_label.config(text="오류 발생!")
        messagebox.showerror("오류", f"변환 중 오류가 발생했습니다: {str(e)}")
        stop_timer = True
    finally:
        is_running = False
        stop_timer = True
        start_button.config(state=tk.NORMAL)
        pdf_button.config(state=tk.NORMAL)
        progress_var.set(0)
        root.after(0, lambda: root.title("PDF to PNG 변환기"))

def update_title(start_time):
    def update():
        global stop_timer
        while not stop_timer:
            elapsed_time = time.time() - start_time
            root.after(0, lambda: root.title(f"PDF to PNG 변환기 - 경과 시간: {elapsed_time:.1f}초"))
            time.sleep(1)
    
    threading.Thread(target=update, daemon=True).start()

def select_pdf_file():
    file_path = filedialog.askopenfilename(title="PDF 파일을 선택하세요", filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(0, file_path)
    else:
        messagebox.showwarning("경고", "파일을 선택하지 않았습니다.")

def start_conversion():
    global is_running
    if is_running:
        return

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

if __name__ == "__main__":
    root = tk.Tk()
    root.title("PDF to PNG 변환기")
    root.geometry("362x350")

    pdf_label = tk.Label(root, text="PDF 파일 선택:")
    pdf_label.grid(row=0, column=0, columnspan=2, pady=10)

    file_frame = tk.Frame(root)
    file_frame.grid(row=1, column=0, columnspan=2, pady=5, padx=5)

    pdf_entry = tk.Entry(file_frame, width=40)
    pdf_entry.grid(row=0, column=0, padx=(0, 5))

    pdf_button = tk.Button(file_frame, text="파일 선택", command=select_pdf_file, background="tomato")
    pdf_button.grid(row=0, column=1)

    progress_label = tk.Label(root, text="대기 중...", font=("Arial", 10))
    progress_label.grid(row=3, column=0,columnspan=2, pady=5)

    progress_var = tk.IntVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=300)
    progress_bar.grid(row=4, column=0, columnspan=2, pady=10)

    start_button = tk.Button(root, text="변환 시작", command=start_conversion, background="steel blue")
    start_button.grid(row=2, column=0, columnspan=2, pady=20)

    # 홍보 이미지 추가
    add_promo_image()

    root.mainloop()
