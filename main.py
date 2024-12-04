import os
import pandas as pd
from docx import Document
import win32com.client  
import fitz  
from PyPDF2 import PdfReader
from collections import Counter
import re

# Fungsi untuk membaca file .txt
def read_txt(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()

# Fungsi untuk membaca file .docx
def read_docx(file_path):
    doc = Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

# Fungsi untuk membaca file .doc (menggunakan pywin32)
def read_doc(file_path):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(file_path)
        full_text = doc.Content.Text
        doc.Close()
        word.Quit()
        return full_text
    except Exception as e:
        print(f"Terjadi kesalahan saat membuka file DOC: {e}")
        return ""

# Fungsi untuk membaca file .pdf
def read_pdf(file_path):
    doc = fitz.open(file_path)
    full_text = []
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        full_text.append(page.get_text())
    return '\n'.join(full_text)

# Fungsi untuk memuat stopwords dari file CSV
def load_stopwords_from_csv(file_path):
    stopwords_df = pd.read_csv(file_path, header=None)
    return set(stopwords_df[0].str.strip().tolist())

# Fungsi untuk memfilter stopwords
def remove_stopwords(text, stop_words):
    # Tokenisasi teks dengan memisahkan berdasarkan spasi dan menghapus karakter non-huruf
    tokens = re.findall(r'\b\w+\b', text.lower())
    # Hapus stopwords
    filtered_words = [word for word in tokens if word not in stop_words]
    return filtered_words

# Fungsi untuk menghitung kata penting
def count_important_words(text, stop_words):
    # Hapus stopwords
    filtered_words = remove_stopwords(text, stop_words)
    # Hitung frekuensi kata
    word_counts = Counter(filtered_words)
    return word_counts

# Fungsi untuk memproses file sesuai format yang dipilih
def process_file(file_path, stopwords):
    print(f"\nMembaca file: {file_path}")
    
    if file_path.endswith('.txt'):
        text = read_txt(file_path)
    elif file_path.endswith('.docx'):
        text = read_docx(file_path)
    elif file_path.endswith('.doc'):
        text = read_doc(file_path)
    elif file_path.endswith('.pdf'):
        text = read_pdf(file_path)
    else:
        print(f"Format file {file_path} tidak didukung.")
        return
    
    word_counts = count_important_words(text, stopwords)
    
    if word_counts:
        print("Jumlah kata penting di file ini:")
        for word, count in word_counts.items():
            print(f"Kata '{word}' ({count} kali).")
    else:
        print("Tidak ada kata penting yang ditemukan.")

# Fungsi utama untuk meminta input dari terminal
def main():
    stopword_file = 'data/stopwordbahasa.csv'

    stopwords = load_stopwords_from_csv(stopword_file)

    print("\nPilih format untuk membaca isi dokumen:")
    print("1. DOCX/DOC")
    print("2. TXT")
    print("3. PDF")
    
    choice = input("Masukkan 1/2/3: ").strip()

    if choice == '1':
        file_format = '.docx'
    elif choice == '2':
        file_format = '.txt'
    elif choice == '3':
        file_format = '.pdf'
    else:
        print("Pilihan tidak valid!")
        return

    file_name = input(f"Masukkan direktori atau dokumen ({file_format}): ").strip()

    if not file_name.endswith(file_format):
        file_name += file_format

    if os.path.isfile(file_name):
        process_file(file_name, stopwords)
    else:
        print(f"File {file_name} tidak ditemukan atau tidak valid.")

if __name__ == "__main__":
    main()
