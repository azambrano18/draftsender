def callback_progreso_gui(fila_actual, total_filas, barra_progreso, porcentaje_var, status_var, status_label, frame_progreso):
    porcentaje = int((fila_actual / total_filas) * 100)
    barra_progreso["value"] = porcentaje
    porcentaje_var.set(f"{porcentaje}%")
    status_var.set(f"Procesando fila {fila_actual} de {total_filas}")
    status_label.pack(side="bottom", pady=(0, 5))
    frame_progreso.pack(side="bottom", fill="x", padx=10, pady=5)
    status_label.update_idletasks()
