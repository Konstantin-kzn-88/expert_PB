import random
import sys
import tkinter as tk
from tkinter import messagebox
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("Ошибка: не установлен pandas.")
    print("Установите зависимости командой:")
    print("pip install pandas openpyxl")
    sys.exit(1)

FILE_NAME = "Д-2.xlsx"
OPTION_LETTERS = ["A", "B", "C", "D", "E", "F"]


class QuizApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Тест")
        self.root.geometry("1000x720")
        self.root.minsize(950, 680)

        self.all_questions = self.load_questions()

        if not self.all_questions:
            messagebox.showerror("Ошибка", "В файле нет вопросов.")
            self.root.destroy()
            return

        self.total_unique_questions = len(self.all_questions)

        self.total_answered = 0
        self.total_correct = 0

        self.current_round_questions = self.all_questions[:]
        random.shuffle(self.current_round_questions)

        self.wrong_questions_next_round = []

        self.round_number = 1
        self.current_index = 0

        self.current_correct_letter = None
        self.current_answers = {}

        self.build_ui()
        self.show_question()

    def load_questions(self):
        file_path = Path(__file__).resolve().parent / FILE_NAME

        if not file_path.exists():
            messagebox.showerror(
                "Ошибка",
                f"Файл '{FILE_NAME}' не найден.\nПоложите его рядом с программой."
            )
            sys.exit(1)

        try:
            df = pd.read_excel(file_path)
        except ImportError:
            messagebox.showerror(
                "Ошибка",
                "Не установлен openpyxl.\nУстановите:\npip install openpyxl"
            )
            sys.exit(1)
        except Exception as e:
            messagebox.showerror("Ошибка чтения Excel", str(e))
            sys.exit(1)

        required_columns = ["question", "correct"]
        for column in required_columns:
            if column not in df.columns:
                messagebox.showerror(
                    "Ошибка",
                    f"В Excel нет обязательной колонки: {column}"
                )
                sys.exit(1)

        questions = df.to_dict(orient="records")
        return questions

    def build_ui(self):
        self.root.configure(padx=12, pady=12)

        top_frame = tk.Frame(self.root)
        top_frame.pack(fill="x", pady=(0, 10))

        self.progress_label = tk.Label(
            top_frame,
            text="Раунд 1 | Вопрос 0/0",
            font=("Arial", 14, "bold")
        )
        self.progress_label.pack(side="left")

        self.stats_label = tk.Label(
            top_frame,
            text="",
            font=("Arial", 12),
            justify="right"
        )
        self.stats_label.pack(side="right")

        self.question_label = tk.Label(
            self.root,
            text="",
            font=("Arial", 15),
            wraplength=950,
            justify="left",
            anchor="w"
        )
        self.question_label.pack(fill="x", pady=(10, 15), anchor="w")

        self.answer_var = tk.StringVar(value="")
        self.buttons = []

        self.answers_frame = tk.Frame(self.root)
        self.answers_frame.pack(fill="x", pady=(0, 15), anchor="w")

        for letter in OPTION_LETTERS:
            rb = tk.Radiobutton(
                self.answers_frame,
                text="",
                variable=self.answer_var,
                value=letter,
                font=("Arial", 13),
                wraplength=950,
                justify="left",
                anchor="w"
            )
            rb.pack(fill="x", anchor="w", pady=3)
            self.buttons.append(rb)

        self.result_title = tk.Label(
            self.root,
            text="",
            font=("Arial", 14, "bold"),
            anchor="w"
        )
        self.result_title.pack(anchor="w")

        self.result_answer = tk.Label(
            self.root,
            text="Выберите вариант и нажмите 'Ответить'",
            font=("Arial", 12),
            wraplength=950,
            justify="left",
            anchor="w"
        )
        self.result_answer.pack(fill="x", pady=(0, 15), anchor="w")

        button_frame = tk.Frame(self.root)
        button_frame.pack(fill="x", pady=(5, 0))

        self.answer_button = tk.Button(
            button_frame,
            text="Ответить",
            command=self.check_answer,
            width=15
        )
        self.answer_button.pack(side="left")

        self.next_button = tk.Button(
            button_frame,
            text="Следующий вопрос",
            command=self.next_question,
            state="disabled",
            width=18
        )
        self.next_button.pack(side="left", padx=10)

        self.exit_button = tk.Button(
            button_frame,
            text="Выход",
            command=self.root.destroy,
            width=10
        )
        self.exit_button.pack(side="right")

    def shuffle_answers(self, question_row):
        answers = []

        for letter in OPTION_LETTERS:
            text = question_row.get(letter)
            if text is not None and str(text) != "nan":
                answers.append((letter, str(text)))

        correct_letter = str(question_row.get("correct", "")).strip().upper()

        correct_text = None
        for letter, text in answers:
            if letter == correct_letter:
                correct_text = text
                break

        random.shuffle(answers)

        new_answers = {}
        for i, (_, text) in enumerate(answers):
            new_letter = OPTION_LETTERS[i]
            new_answers[new_letter] = text

        new_correct_letter = None
        for letter, text in new_answers.items():
            if text == correct_text:
                new_correct_letter = letter
                break

        return new_answers, new_correct_letter

    def show_question(self):
        if self.current_index >= len(self.current_round_questions):
            self.finish_round()
            return

        question = self.current_round_questions[self.current_index]

        self.current_answers, self.current_correct_letter = self.shuffle_answers(question)

        self.answer_var.set("")
        self.result_title.config(text="", fg="black")
        self.result_answer.config(text="Выберите вариант и нажмите 'Ответить'")
        self.answer_button.config(state="normal")
        self.next_button.config(state="disabled")

        self.progress_label.config(
            text=(
                f"Раунд {self.round_number} | "
                f"Вопрос {self.current_index + 1}/{len(self.current_round_questions)}"
            )
        )

        self.question_label.config(text=str(question.get("question", "")))

        answer_items = list(self.current_answers.items())

        for i, rb in enumerate(self.buttons):
            if i < len(answer_items):
                letter, text = answer_items[i]
                rb.config(
                    text=f"{letter}: {text}",
                    value=letter,
                    state="normal"
                )
                rb.pack(fill="x", anchor="w", pady=3)
            else:
                rb.pack_forget()

        self.update_stats()

    def check_answer(self):
        user_answer = self.answer_var.get()

        if not user_answer:
            messagebox.showwarning("Внимание", "Сначала выберите вариант ответа.")
            return

        question = self.current_round_questions[self.current_index]
        correct_text = self.current_answers[self.current_correct_letter]

        self.total_answered += 1

        if user_answer == self.current_correct_letter:
            self.total_correct += 1
            self.result_title.config(
                text="Верно",
                fg="green"
            )
            self.result_answer.config(
                text=f"Правильный ответ: {self.current_correct_letter}\n{correct_text}"
            )
        else:
            self.wrong_questions_next_round.append(question)
            self.result_title.config(
                text="Неверно",
                fg="red"
            )
            self.result_answer.config(
                text=(
                    f"Правильный ответ: {self.current_correct_letter}\n"
                    f"{correct_text}\n\n"
                    f"Этот вопрос будет повторен в следующем раунде."
                )
            )

        self.answer_button.config(state="disabled")
        self.next_button.config(state="normal")
        self.update_stats()

    def next_question(self):
        self.current_index += 1
        self.show_question()

    def finish_round(self):
        wrong_count = len(self.wrong_questions_next_round)

        if wrong_count == 0:
            self.finish_test()
            return

        messagebox.showinfo(
            "Раунд завершен",
            f"Раунд {self.round_number} завершен.\n\n"
            f"Ошибок в этом раунде: {wrong_count}\n"
            f"Сейчас начнется повтор только этих вопросов."
        )

        self.current_round_questions = self.wrong_questions_next_round[:]
        random.shuffle(self.current_round_questions)

        self.wrong_questions_next_round = []
        self.round_number += 1
        self.current_index = 0

        self.show_question()

    def update_stats(self):
        percent = (self.total_correct / self.total_answered * 100) if self.total_answered else 0.0
        remaining_to_master = len(self.wrong_questions_next_round)

        if self.current_index < len(self.current_round_questions):
            current_round_remaining = len(self.current_round_questions) - self.current_index
        else:
            current_round_remaining = 0

        self.stats_label.config(
            text=(
                f"Уникальных вопросов: {self.total_unique_questions}\n"
                f"Всего ответов: {self.total_answered} | "
                f"Верных: {self.total_correct} | "
                f"Процент: {percent:.2f}%\n"
                f"Ошибок на повтор в этом раунде: {remaining_to_master} | "
                f"Осталось в текущем раунде: {current_round_remaining}"
            )
        )

    def finish_test(self):
        percent = (self.total_correct / self.total_answered * 100) if self.total_answered else 0.0

        messagebox.showinfo(
            "Тест завершен",
            "Поздравляю, достигнуто 100% по всем вопросам.\n\n"
            f"Уникальных вопросов: {self.total_unique_questions}\n"
            f"Всего данных ответов: {self.total_answered}\n"
            f"Верных ответов: {self.total_correct}\n"
            f"Общий процент верных ответов: {percent:.2f}%"
        )

        self.root.destroy()


def main():
    root = tk.Tk()
    QuizApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()