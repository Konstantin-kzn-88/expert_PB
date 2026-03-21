import random
import sys
import json
import tkinter as tk
from tkinter import messagebox
from pathlib import Path
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("Ошибка: не установлен pandas.")
    print("Установите зависимости командой:")
    print("pip install pandas openpyxl")
    sys.exit(1)

OPTION_LETTERS = ["A", "B", "C", "D", "E", "F"]
STATS_FILE = "quiz_stats.json"


# ─────────────────────────── Statistics helpers ───────────────────────────────

def load_stats() -> dict:
    path = Path(__file__).resolve().parent / STATS_FILE
    if path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_stats(stats: dict):
    path = Path(__file__).resolve().parent / STATS_FILE
    with open(path, "w", encoding="utf-8") as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)


def record_session(file_name: str, total_answered: int, total_correct: int, rounds: int):
    stats = load_stats()
    if file_name not in stats:
        stats[file_name] = []
    stats[file_name].append({
        "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "answered": total_answered,
        "correct": total_correct,
        "percent": round(total_correct / total_answered * 100, 1) if total_answered else 0.0,
        "rounds": rounds,
    })
    save_stats(stats)


# ─────────────────────────── Stats window ────────────────────────────────────

class StatsWindow:
    PAD = 40
    BAR_W = 36
    COLORS = {
        "bg":      "#f5f5f5",
        "card":    "#ffffff",
        "accent":  "#4a90d9",
        "good":    "#4caf50",
        "warn":    "#ff9800",
        "bad":     "#f44336",
        "text":    "#212121",
        "subtext": "#757575",
        "grid":    "#e0e0e0",
    }

    def __init__(self, parent):
        self.win = tk.Toplevel(parent)
        self.win.title("Статистика")
        self.win.geometry("860x640")
        self.win.minsize(700, 520)
        self.win.configure(bg=self.COLORS["bg"])

        self.stats = load_stats()
        # only session keys (not __questions keys)
        self.file_names = sorted(k for k in self.stats if not k.endswith("__questions"))
        self.selected_file = tk.StringVar(
            value=self.file_names[0] if self.file_names else ""
        )
        self._build()

    def _build(self):
        if not self.file_names:
            tk.Label(
                self.win,
                text="Нет данных.\nПройдите хотя бы один тест до конца.",
                font=("Arial", 14),
                bg=self.COLORS["bg"],
                fg=self.COLORS["subtext"],
            ).pack(expand=True)
            return

        # file selector
        top = tk.Frame(self.win, bg=self.COLORS["bg"])
        top.pack(fill="x", padx=16, pady=(12, 4))

        tk.Label(top, text="Файл:", font=("Arial", 11),
                 bg=self.COLORS["bg"]).pack(side="left")
        menu = tk.OptionMenu(top, self.selected_file, *self.file_names,
                             command=lambda _: self._refresh())
        menu.config(font=("Arial", 11), width=22)
        menu.pack(side="left", padx=8)
        tk.Button(top, text="Очистить историю этого файла",
                  command=self._clear_file, font=("Arial", 10)).pack(side="right")

        # summary cards
        self.cards_frame = tk.Frame(self.win, bg=self.COLORS["bg"])
        self.cards_frame.pack(fill="x", padx=16, pady=6)

        # chart
        tk.Label(self.win, text="% правильных по сессиям",
                 font=("Arial", 11, "bold"),
                 bg=self.COLORS["bg"]).pack(anchor="w", padx=16)
        self.chart_canvas = tk.Canvas(self.win, bg=self.COLORS["card"],
                                      highlightthickness=1,
                                      highlightbackground=self.COLORS["grid"],
                                      height=220)
        self.chart_canvas.pack(fill="x", padx=16, pady=(2, 8))
        self.chart_canvas.bind("<Configure>", lambda e: self._redraw_chart())

        # weak questions
        tk.Label(self.win, text="Слабые вопросы (чаще всего ошибаюсь)",
                 font=("Arial", 11, "bold"),
                 bg=self.COLORS["bg"]).pack(anchor="w", padx=16)
        self.weak_frame = tk.Frame(self.win, bg=self.COLORS["bg"])
        self.weak_frame.pack(fill="both", expand=True, padx=16, pady=(2, 12))

        self._refresh()

    def _refresh(self):
        fname = self.selected_file.get()
        sessions = self.stats.get(fname, [])
        self._draw_cards(sessions)
        self._redraw_chart()
        self._draw_weak(fname)

    def _redraw_chart(self):
        fname = self.selected_file.get()
        sessions = self.stats.get(fname, [])
        self.chart_canvas.update_idletasks()
        self._draw_chart(sessions)

    def _draw_cards(self, sessions):
        for w in self.cards_frame.winfo_children():
            w.destroy()
        if not sessions:
            return

        percents = [s["percent"] for s in sessions]
        last_pct = percents[-1]
        best_pct = max(percents)
        total_ans = sum(s["answered"] for s in sessions)
        total_cor = sum(s["correct"] for s in sessions)
        overall = round(total_cor / total_ans * 100, 1) if total_ans else 0
        avg_rounds = round(sum(s["rounds"] for s in sessions) / len(sessions), 1)
        trend = percents[-1] - percents[-2] if len(percents) >= 2 else None

        cards = [
            ("Последняя\nсессия",  f"{last_pct}%",  self._pct_color(last_pct)),
            ("Лучший\nрезультат", f"{best_pct}%",  self._pct_color(best_pct)),
            ("Общий %",            f"{overall}%",   self._pct_color(overall)),
            ("Средних\nраундов",   str(avg_rounds), self.COLORS["accent"]),
            ("Сессий\nпройдено",   str(len(sessions)), self.COLORS["accent"]),
            ("Тренд",
             (f"+{trend:.1f}%" if trend and trend >= 0
              else f"{trend:.1f}%" if trend is not None else "—"),
             (self.COLORS["good"] if trend and trend >= 0
              else self.COLORS["bad"] if trend is not None
              else self.COLORS["subtext"])),
        ]

        for i, (label, value, color) in enumerate(cards):
            card = tk.Frame(self.cards_frame, bg=self.COLORS["card"],
                            highlightthickness=1,
                            highlightbackground=self.COLORS["grid"])
            card.grid(row=0, column=i, padx=5, pady=4, sticky="nsew")
            self.cards_frame.columnconfigure(i, weight=1)
            tk.Label(card, text=value, font=("Arial", 17, "bold"),
                     fg=color, bg=self.COLORS["card"]).pack(pady=(8, 0))
            tk.Label(card, text=label, font=("Arial", 8),
                     fg=self.COLORS["subtext"],
                     bg=self.COLORS["card"]).pack(pady=(0, 8))

    def _draw_chart(self, sessions):
        c = self.chart_canvas
        c.delete("all")
        W = c.winfo_width() or 800
        H = 220
        pad = self.PAD

        if not sessions:
            c.create_text(W // 2, H // 2, text="Нет данных",
                          fill=self.COLORS["subtext"], font=("Arial", 12))
            return

        for pct in (0, 25, 50, 75, 100):
            y = pad + (H - 2 * pad) * (1 - pct / 100)
            c.create_line(pad, y, W - pad // 2, y,
                          fill=self.COLORS["grid"], dash=(4, 4))
            c.create_text(pad - 6, y, text=f"{pct}%",
                          anchor="e", font=("Arial", 8),
                          fill=self.COLORS["subtext"])

        percents = [s["percent"] for s in sessions]
        n = len(percents)
        slot_w = max(self.BAR_W + 6, (W - 2 * pad) // max(n, 1))
        bar_w = min(self.BAR_W, slot_w - 4)

        for i, (s, pct) in enumerate(zip(sessions, percents)):
            x_center = pad + slot_w * i + slot_w // 2
            bar_h = (H - 2 * pad) * pct / 100
            x0 = x_center - bar_w // 2
            x1 = x_center + bar_w // 2
            y0 = H - pad - bar_h
            y1 = H - pad
            color = self._pct_color(pct)
            c.create_rectangle(x0, y0, x1, y1, fill=color, outline="")
            c.create_text(x_center, y0 - 4, text=f"{pct:.0f}%",
                          font=("Arial", 7, "bold"), fill=color, anchor="s")
            date_short = s["date"][5:10]
            c.create_text(x_center, H - pad + 4, text=date_short,
                          font=("Arial", 7), fill=self.COLORS["subtext"], anchor="n")

        if n >= 2:
            points = []
            for i, pct in enumerate(percents):
                x = pad + slot_w * i + slot_w // 2
                y = pad + (H - 2 * pad) * (1 - pct / 100)
                points.append((x, y))
            for i in range(len(points) - 1):
                c.create_line(points[i][0], points[i][1],
                              points[i + 1][0], points[i + 1][1],
                              fill=self.COLORS["text"], width=2)
            for x, y in points:
                c.create_oval(x - 3, y - 3, x + 3, y + 3,
                              fill=self.COLORS["text"], outline="")

    def _draw_weak(self, fname):
        for w in self.weak_frame.winfo_children():
            w.destroy()

        q_key = fname + "__questions"
        q_stats = load_stats().get(q_key, {})

        if not q_stats:
            tk.Label(self.weak_frame,
                     text="Данные по отдельным вопросам появятся после следующего теста.",
                     font=("Arial", 10), bg=self.COLORS["bg"],
                     fg=self.COLORS["subtext"]).pack(anchor="w")
            return

        sorted_q = sorted(q_stats.items(),
                          key=lambda x: x[1]["wrong"], reverse=True)[:8]

        cols = ("Вопрос", "Ошибок", "Попыток", "% ошибок")
        col_w = (54, 7, 7, 8)

        header = tk.Frame(self.weak_frame, bg=self.COLORS["grid"])
        header.pack(fill="x")
        for col, cw in zip(cols, col_w):
            tk.Label(header, text=col, font=("Arial", 9, "bold"),
                     bg=self.COLORS["grid"], width=cw,
                     anchor="w").pack(side="left", padx=4)

        for i, (q_text, data) in enumerate(sorted_q):
            bg = self.COLORS["card"] if i % 2 == 0 else self.COLORS["bg"]
            row = tk.Frame(self.weak_frame, bg=bg)
            row.pack(fill="x")
            wrong = data["wrong"]
            attempts = data["attempts"]
            pct_wrong = round(wrong / attempts * 100) if attempts else 0
            short_q = q_text[:72] + "…" if len(q_text) > 72 else q_text
            for val, cw in zip(
                (short_q, str(wrong), str(attempts), f"{pct_wrong}%"), col_w
            ):
                tk.Label(row, text=val, font=("Arial", 9),
                         bg=bg, anchor="w", width=cw).pack(side="left", padx=4, pady=1)

    def _pct_color(self, pct):
        if pct >= 80:
            return self.COLORS["good"]
        if pct >= 50:
            return self.COLORS["warn"]
        return self.COLORS["bad"]

    def _clear_file(self):
        fname = self.selected_file.get()
        if not fname:
            return
        if not messagebox.askyesno("Подтверждение",
                                   f"Удалить всю историю для «{fname}»?"):
            return
        stats = load_stats()
        stats.pop(fname, None)
        stats.pop(fname + "__questions", None)
        save_stats(stats)
        self.stats = stats
        self.file_names = sorted(k for k in stats if not k.endswith("__questions"))
        if self.file_names:
            self.selected_file.set(self.file_names[0])
            self._refresh()
        else:
            self.win.destroy()


# ─────────────────────────── File selector ───────────────────────────────────

class FileSelectApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Выбор теста")
        self.root.resizable(False, False)
        self.build_ui()

    def build_ui(self):
        self.root.configure(padx=20, pady=20)

        tk.Label(self.root, text="Выберите файл с вопросами:",
                 font=("Arial", 14, "bold")).pack(pady=(0, 14))

        base_dir = Path(__file__).resolve().parent
        xlsx_files = sorted(base_dir.glob("*.xlsx"))

        if not xlsx_files:
            tk.Label(self.root, text="Нет .xlsx файлов рядом с программой.",
                     font=("Arial", 12), fg="red").pack()
        else:
            btn_frame = tk.Frame(self.root)
            btn_frame.pack()
            for xlsx in xlsx_files:
                tk.Button(btn_frame, text=xlsx.name, font=("Arial", 13),
                          width=28, command=lambda f=xlsx: self.launch(f)
                          ).pack(fill="x", pady=4)

        bottom = tk.Frame(self.root)
        bottom.pack(pady=(18, 0), fill="x")
        tk.Button(bottom, text="📊 Статистика", command=self._open_stats,
                  font=("Arial", 11), width=16).pack(side="left")
        tk.Button(bottom, text="Выход", command=self.root.destroy,
                  width=10, font=("Arial", 11)).pack(side="right")

    def _open_stats(self):
        StatsWindow(self.root)

    def launch(self, file_path):
        self.root.destroy()
        root2 = tk.Tk()
        QuizApp(root2, file_path)
        root2.mainloop()


# ─────────────────────────── Quiz ────────────────────────────────────────────

class QuizApp:
    def __init__(self, root, file_path: Path):
        self.root = root
        self.file_path = file_path
        self.root.title(f"Тест — {file_path.name}")
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
        self.round_number = 1
        self.current_index = 0
        self.current_correct_letter = None
        self.current_answers = {}

        self.current_round_questions = self.all_questions[:]
        random.shuffle(self.current_round_questions)
        self.wrong_questions_next_round = []

        # {question_text: {"attempts": n, "wrong": n}}
        self.question_stats: dict = {}

        self.build_ui()
        self.show_question()

    def load_questions(self):
        if not self.file_path.exists():
            messagebox.showerror("Ошибка",
                                 f"Файл '{self.file_path.name}' не найден.")
            sys.exit(1)
        try:
            df = pd.read_excel(self.file_path)
        except ImportError:
            messagebox.showerror("Ошибка",
                                 "Не установлен openpyxl.\npip install openpyxl")
            sys.exit(1)
        except Exception as e:
            messagebox.showerror("Ошибка чтения Excel", str(e))
            sys.exit(1)

        for col in ("question", "correct"):
            if col not in df.columns:
                messagebox.showerror("Ошибка",
                                     f"В Excel нет обязательной колонки: {col}")
                sys.exit(1)
        return df.to_dict(orient="records")

    def build_ui(self):
        self.root.configure(padx=12, pady=12)

        top_frame = tk.Frame(self.root)
        top_frame.pack(fill="x", pady=(0, 10))
        self.progress_label = tk.Label(top_frame, text="Раунд 1 | Вопрос 0/0",
                                       font=("Arial", 14, "bold"))
        self.progress_label.pack(side="left")
        self.stats_label = tk.Label(top_frame, text="", font=("Arial", 12),
                                    justify="right")
        self.stats_label.pack(side="right")

        self.question_label = tk.Label(self.root, text="", font=("Arial", 15),
                                       wraplength=950, justify="left", anchor="w")
        self.question_label.pack(fill="x", pady=(10, 15), anchor="w")

        self.answer_var = tk.StringVar(value="")
        self.buttons = []
        self.answers_frame = tk.Frame(self.root)
        self.answers_frame.pack(fill="x", pady=(0, 15), anchor="w")
        for letter in OPTION_LETTERS:
            rb = tk.Radiobutton(self.answers_frame, text="",
                                variable=self.answer_var, value=letter,
                                font=("Arial", 13), wraplength=950,
                                justify="left", anchor="w")
            rb.pack(fill="x", anchor="w", pady=3)
            self.buttons.append(rb)

        self.result_title = tk.Label(self.root, text="",
                                     font=("Arial", 14, "bold"), anchor="w")
        self.result_title.pack(anchor="w")
        self.result_answer = tk.Label(
            self.root, text="Выберите вариант и нажмите 'Ответить'",
            font=("Arial", 12), wraplength=950, justify="left", anchor="w")
        self.result_answer.pack(fill="x", pady=(0, 15), anchor="w")

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill="x", pady=(5, 0))
        self.answer_button = tk.Button(btn_frame, text="Ответить",
                                       command=self.check_answer, width=15)
        self.answer_button.pack(side="left")
        self.next_button = tk.Button(btn_frame, text="Следующий вопрос",
                                     command=self.next_question,
                                     state="disabled", width=18)
        self.next_button.pack(side="left", padx=10)
        tk.Button(btn_frame, text="← Сменить файл",
                  command=self.back_to_select, width=16).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Выход",
                  command=self.root.destroy, width=10).pack(side="right")

    def shuffle_answers(self, question_row):
        answers = [(l, str(question_row[l]))
                   for l in OPTION_LETTERS
                   if question_row.get(l) is not None
                   and str(question_row.get(l)) != "nan"]
        correct_letter = str(question_row.get("correct", "")).strip().upper()
        correct_text = next((t for l, t in answers if l == correct_letter), None)
        random.shuffle(answers)
        new_answers = {OPTION_LETTERS[i]: t for i, (_, t) in enumerate(answers)}
        new_correct = next((l for l, t in new_answers.items()
                            if t == correct_text), None)
        return new_answers, new_correct

    def show_question(self):
        if self.current_index >= len(self.current_round_questions):
            self.finish_round()
            return

        question = self.current_round_questions[self.current_index]
        self.current_answers, self.current_correct_letter = \
            self.shuffle_answers(question)

        self.answer_var.set("")
        self.result_title.config(text="", fg="black")
        self.result_answer.config(text="Выберите вариант и нажмите 'Ответить'")
        self.answer_button.config(state="normal")
        self.next_button.config(state="disabled")

        self.progress_label.config(
            text=(f"Раунд {self.round_number} | "
                  f"Вопрос {self.current_index + 1}/"
                  f"{len(self.current_round_questions)}")
        )
        self.question_label.config(text=str(question.get("question", "")))

        answer_items = list(self.current_answers.items())
        for i, rb in enumerate(self.buttons):
            if i < len(answer_items):
                letter, text = answer_items[i]
                rb.config(text=f"{letter}: {text}", value=letter, state="normal")
                rb.pack(fill="x", anchor="w", pady=3)
            else:
                rb.pack_forget()

        self.update_stats_label()

    def check_answer(self):
        user_answer = self.answer_var.get()
        if not user_answer:
            messagebox.showwarning("Внимание", "Сначала выберите вариант ответа.")
            return

        question = self.current_round_questions[self.current_index]
        correct_text = self.current_answers[self.current_correct_letter]
        q_text = str(question.get("question", ""))

        self.total_answered += 1
        qs = self.question_stats.setdefault(q_text, {"attempts": 0, "wrong": 0})
        qs["attempts"] += 1

        if user_answer == self.current_correct_letter:
            self.total_correct += 1
            self.result_title.config(text="Верно", fg="green")
            self.result_answer.config(
                text=f"Правильный ответ: {self.current_correct_letter}\n{correct_text}")
        else:
            qs["wrong"] += 1
            self.wrong_questions_next_round.append(question)
            self.result_title.config(text="Неверно", fg="red")
            self.result_answer.config(
                text=(f"Правильный ответ: {self.current_correct_letter}\n"
                      f"{correct_text}\n\n"
                      f"Этот вопрос будет повторен в следующем раунде."))

        self.answer_button.config(state="disabled")
        self.next_button.config(state="normal")
        self.update_stats_label()

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
            f"Сейчас начнется повтор только этих вопросов.")
        self.current_round_questions = self.wrong_questions_next_round[:]
        random.shuffle(self.current_round_questions)
        self.wrong_questions_next_round = []
        self.round_number += 1
        self.current_index = 0
        self.show_question()

    def update_stats_label(self):
        percent = (self.total_correct / self.total_answered * 100
                   if self.total_answered else 0.0)
        remaining = len(self.wrong_questions_next_round)
        left = (len(self.current_round_questions) - self.current_index
                if self.current_index < len(self.current_round_questions) else 0)
        self.stats_label.config(
            text=(f"Уникальных вопросов: {self.total_unique_questions}\n"
                  f"Всего ответов: {self.total_answered} | "
                  f"Верных: {self.total_correct} | "
                  f"Процент: {percent:.2f}%\n"
                  f"Ошибок на повтор в этом раунде: {remaining} | "
                  f"Осталось в текущем раунде: {left}"))

    def finish_test(self):
        percent = (self.total_correct / self.total_answered * 100
                   if self.total_answered else 0.0)

        record_session(
            file_name=self.file_path.name,
            total_answered=self.total_answered,
            total_correct=self.total_correct,
            rounds=self.round_number,
        )

        # merge per-question stats
        all_stats = load_stats()
        q_key = self.file_path.name + "__questions"
        saved_q = all_stats.setdefault(q_key, {})
        for q_text, data in self.question_stats.items():
            entry = saved_q.setdefault(q_text, {"attempts": 0, "wrong": 0})
            entry["attempts"] += data["attempts"]
            entry["wrong"] += data["wrong"]
        save_stats(all_stats)

        messagebox.showinfo(
            "Тест завершен",
            "Поздравляю, достигнуто 100% по всем вопросам.\n\n"
            f"Уникальных вопросов: {self.total_unique_questions}\n"
            f"Всего данных ответов: {self.total_answered}\n"
            f"Верных ответов: {self.total_correct}\n"
            f"Общий процент верных ответов: {percent:.2f}%")
        self.root.destroy()

    def back_to_select(self):
        self.root.destroy()
        root2 = tk.Tk()
        FileSelectApp(root2)
        root2.mainloop()


# ─────────────────────────── Entry point ─────────────────────────────────────

def main():
    root = tk.Tk()
    FileSelectApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
