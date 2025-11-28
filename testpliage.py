import tkinter as tk

class FoldSelectionWindow:
    def __init__(self, master):
        self.master = master
        master.title("Sélection du pliage")
        master.geometry("400x300")

        self.options = [
            ("Pliage en 2", self.get_fold_lines(2)),
            ("Pliage en 4", self.get_fold_lines(4)),
            ("Pliage en 8", self.get_fold_lines(8))
        ]

        self.selected_option = tk.StringVar()
        self.selected_option.set(self.options[0][0])

        for i, (text, lines) in enumerate(self.options):
            frame = tk.Frame(master)
            frame.pack(pady=8, anchor='w')

            # Case à cocher
            rb = tk.Radiobutton(frame, text=text, variable=self.selected_option, value=text)
            rb.pack(side='left')

            # Canvas pour le visuel
            canvas = tk.Canvas(frame, width=120, height=80, bg='white')
            canvas.pack(side='left', padx=10)
            self.draw_fold(canvas, lines)

    def get_fold_lines(self, fold_type):
        if fold_type == 2:
            # Pliage en deux = une ligne verticale au centre
            return [(0.5, 0, 0.5, 1)]
        elif fold_type == 4:
            # Pliage en quatre = 1 verticale + 1 horizontale
            return [(0.5, 0, 0.5, 1), (0, 0.5, 1, 0.5)]
        elif fold_type == 8:
            # Pliage en huit = 3 verticales + 1 horizontale
            return [(1/4, 0, 1/4, 1), (1/2, 0, 1/2, 1), (3/4, 0, 3/4, 1), (0, 0.5, 1, 0.5)]
        return []

    def draw_fold(self, canvas, lines):
        w = canvas.winfo_reqwidth()
        h = canvas.winfo_reqheight()
        margin = 6

        # Rectangle représentant la feuille
        canvas.create_rectangle(margin, margin, w-margin, h-margin, outline='black', width=2)

        # Lignes de pliage pointillées
        for x0, y0, x1, y1 in lines:
            canvas.create_line(x0*w, y0*h, x1*w, y1*h, dash=(4, 3), fill='blue', width=1.5)

if __name__ == "__main__":
    root = tk.Tk()
    FoldSelectionWindow(root)
    root.mainloop()
