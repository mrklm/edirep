import tkinter as tk
from tkinter import Toplevel

def show_print_tutorial():
    # Fenêtre compacte
    popup = Toplevel()
    popup.title("Tutoriel d'impression du livret")
    popup.geometry("650x600")  # réduit la taille au strict minimum

    canvas_width = 650
    canvas_height = 600
    canvas = tk.Canvas(popup, width=canvas_width, height=canvas_height, bg="white")
    canvas.pack(fill="both", expand=True)

    cols = 2
    page_h = 80
    page_w = int(page_h * 1.414)
    spacing_x = 70
    spacing_y = 40
    margin_y = 30
    shadow_offset_x = 4
    shadow_offset_y = 4

    # Numéros et textes
    page_numbers = [1, 2, 2, 3, 3, 4, 4]
    page_labels = [
        "1 Recto uniquement",
        "2 Recto",
        "2 Verso",
        "3 Recto",
        "3 Verso",
        "4 Recto",
        "4 Verso",
    ]

    # Calcul du départ horizontal
    total_width = cols * page_w + (cols - 1) * spacing_x
    start_x = (canvas_width - total_width) / 2

    # Dessin des mini-A4
    positions = []
    for i in range(7):
        r = i // cols
        c = i % cols

        x1 = start_x + c * (page_w + spacing_x)
        y1 = margin_y + r * (page_h + spacing_y)
        x2 = x1 + page_w
        y2 = y1 + page_h

        positions.append((x1, y1, x2, y2))

        # Ombre
        canvas.create_rectangle(
            x1 + shadow_offset_x, y1 + shadow_offset_y,
            x2 + shadow_offset_x, y2 + shadow_offset_y,
            fill="#dddddd", outline=""
        )

        # Feuille avec coin coupé
        cut = 15
        if "Verso" in page_labels[i]:
            canvas.create_polygon(
                x1, y1,
                x2, y1,
                x2, y2,
                x1 + cut, y2,
                x1, y2 - cut,
                outline="black",
                fill="white"
            )
        else:
            canvas.create_polygon(
                x1 + cut, y1,
                x2, y1,
                x2, y2,
                x1, y2,
                x1, y1 + cut,
                outline="black",
                fill="white"
            )

        # Numéro centré
        canvas.create_text(
            (x1 + x2)/2,
            (y1 + y2)/2,
            text=str(page_numbers[i]),
            font=("Arial", 18, "bold"),
            fill="black"
        )

        # Annotation
        canvas.create_text(
            (x1 + x2)/2, y2 - 8,
            text=page_labels[i],
            font=("Arial", 9),
            anchor="s",
            fill="black"
        )

    # Fonction pour dessiner le livret avec écart entre pages et ombrage
    def draw_booklet(base_x, base_y, page_count=7, scale=0.5, height_multiplier=3, page_gap=3):
        fold_offset = 10 * scale
        page_half_w = (page_w / 2) * scale
        page_half_h = (page_h / 2) * scale * height_multiplier

        for i in range(page_count):
            dx = i * page_gap  # écart horizontal
            dy = i * page_gap  # écart vertical pour superposition

            # Ombre légère
            canvas.create_rectangle(
                base_x + dx + 2, base_y + dy + 2,
                base_x + dx + 2 + page_half_w*2,
                base_y + dy + 2 + page_half_h,
                fill="#cccccc", outline=""
            )

            # Partie gauche
            canvas.create_polygon(
                base_x + dx, base_y + dy,
                base_x + dx + page_half_w, base_y + dy - fold_offset,
                base_x + dx + page_half_w, base_y + dy + page_half_h - fold_offset,
                base_x + dx, base_y + dy + page_half_h,
                fill="white",
                outline="black"
            )
            # Partie droite
            canvas.create_polygon(
                base_x + dx + page_half_w, base_y + dy - fold_offset,
                base_x + dx + 2*page_half_w, base_y + dy,
                base_x + dx + 2*page_half_w, base_y + dy + page_half_h,
                base_x + dx + page_half_w, base_y + dy + page_half_h - fold_offset,
                fill="white",
                outline="black"
            )

    # Positionner le livret : centré entre 4 Recto et 4 Verso
    ref_x1, ref_y1, ref_x2, ref_y2 = positions[5]  # 4 Recto
    ref_x1v, ref_y1v, ref_x2v, ref_y2v = positions[6]  # 4 Verso

    # Centre horizontal sous 4 Recto et centre vertical à droite de 4 Verso
    center_x = (ref_x1 + ref_x2v) / 2 + 10  # décalage à droite
    center_y = (ref_y2 + ref_y1v + page_h)/2

    draw_booklet(center_x, center_y, page_count=7, scale=0.5, height_multiplier=3, page_gap=4)

# Test rapide
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Programme principal")
    tk.Button(root, text="Afficher tutoriel", command=show_print_tutorial).pack(pady=20)
    root.mainloop()
