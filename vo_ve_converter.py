import tkinter as tk
from tkinter import messagebox
import os
import sys


class TableEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("VO to VE Converter - Adrenaline Driven")
        self.root.geometry("1920x1080")
        self.root.resizable(True, True)

        # Set window icon
        self.set_icon()

        # 8x12 table settings
        self.rows = 8
        self.cols = 12

        # 16x16 table settings
        self.new_rows = 16
        self.new_cols = 16

        # Storage for 8x12 table
        self.x_labels = []
        self.y_labels = []
        self.cells = []

        # Storage for 16x16 extended table (VO)
        self.new_x_labels = []
        self.new_y_labels = []
        self.new_cells = []

        # Storage for 16x16 VE table
        self.ve_x_labels = []
        self.ve_y_labels = []
        self.ve_cells = []

        # User inputs
        self.displacement_var = tk.StringVar(value="1.6")
        self.iat_var = tk.StringVar(value="20")

        # Grids and axes used for calculations
        self.vo_grid = None
        self.new_x = None
        self.new_y = None

        # Selection tracking
        self.selection_start = None
        self.selected_cells = []
        self.current_table = None

        self.create_ui()
        self.bind_events()

    def set_icon(self):
        """Set window icon from logo.ico located near exe/py."""
        try:
            base_paths = []

            if getattr(sys, 'frozen', False):
                # PyInstaller onefile: temp folder + exe folder
                base_paths.append(sys._MEIPASS)
                base_paths.append(os.path.dirname(sys.executable))
            else:
                # Normal .py execution
                base_paths.append(os.path.dirname(os.path.abspath(__file__)))

            # Also check current working directory
            base_paths.append(os.getcwd())

            for base in base_paths:
                icon_path = os.path.join(base, 'logo.ico')
                if os.path.exists(icon_path):
                    self.root.iconbitmap(icon_path)
                    break
        except Exception:
            # If icon fails, ignore – app should still run
            pass

    def create_ui(self):
        # ===== MAIN CONTAINER =====
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # ===== TOP ROW: 8x12 TABLE + PARAMS =====
        top_frame = tk.Frame(main_frame)
        top_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))

        # ----- 8x12 VO TABLE -----
        frame_8x12 = tk.LabelFrame(top_frame, text="Original 8x12 Table (VO)")
        frame_8x12.pack(side=tk.LEFT, padx=(0, 20))

        # X-axis labels (columns)
        for col in range(self.cols):
            entry = tk.Entry(frame_8x12, width=6, bg="#d0d0d0", font=('Arial', 8))
            entry.grid(row=0, column=col + 1, padx=1, pady=1)
            entry.insert(0, str(col + 1))
            self.bind_axis_navigation(entry, self.x_labels, "x", col, "8x12")
            self.x_labels.append(entry)

        # Y-axis labels (rows) + data cells
        for row in range(self.rows):
            label_entry = tk.Entry(frame_8x12, width=6, bg="#d0d0d0", font=('Arial', 8))
            label_entry.grid(row=row + 1, column=0, padx=1, pady=1)
            label_entry.insert(0, str(row + 1))
            self.bind_axis_navigation(label_entry, self.y_labels, "y", row, "8x12")
            self.y_labels.append(label_entry)

            row_cells = []
            for col in range(self.cols):
                entry = tk.Entry(frame_8x12, width=6, font=('Arial', 8))
                entry.grid(row=row + 1, column=col + 1, padx=1, pady=1)
                self.bind_cell_events(entry, self.cells, row, col, "8x12")
                row_cells.append(entry)
            self.cells.append(row_cells)

        # Buttons for 8x12 table
        btn_frame_8x12 = tk.Frame(frame_8x12)
        btn_frame_8x12.grid(row=self.rows + 2, column=0, columnspan=self.cols + 1, pady=5)
        tk.Button(btn_frame_8x12, text="Clear All", command=self.clear_all).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame_8x12, text="Generate 16x16", command=self.generate_16x16).pack(side=tk.LEFT, padx=5)

        # ----- PARAMETERS FRAME -----
        input_frame = tk.LabelFrame(top_frame, text="Parameters")
        input_frame.pack(side=tk.LEFT, padx=20, fill=tk.Y)

        tk.Label(input_frame, text="Displacement (dm³):").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        tk.Entry(input_frame, textvariable=self.displacement_var, width=10).grid(row=0, column=1, padx=5, pady=5)

        tk.Label(input_frame, text="IAT (°C):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        tk.Entry(input_frame, textvariable=self.iat_var, width=10).grid(row=1, column=1, padx=5, pady=5)

        tk.Button(input_frame, text="Calculate VE", command=self.calculate_ve).grid(
            row=2, column=0, columnspan=2, pady=10
        )

        # ===== BOTTOM ROW: 16x16 VO + 16x16 VE =====
        bottom_frame = tk.Frame(main_frame)
        bottom_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # ----- 16x16 VO TABLE -----
        frame_16x16 = tk.LabelFrame(bottom_frame, text="Extended 16x16 Table (VO)")
        frame_16x16.pack(side=tk.LEFT, padx=(0, 10), fill=tk.BOTH, expand=True)

        for col in range(self.new_cols):
            entry = tk.Entry(frame_16x16, width=5, bg="#d0d0d0", font=('Arial', 8))
            entry.grid(row=0, column=col + 1, padx=1, pady=1)
            entry.insert(0, "")
            self.bind_axis_navigation(entry, self.new_x_labels, "x", col, "16x16_vo")
            self.new_x_labels.append(entry)

        for row in range(self.new_rows):
            label_entry = tk.Entry(frame_16x16, width=5, bg="#d0d0d0", font=('Arial', 8))
            label_entry.grid(row=row + 1, column=0, padx=1, pady=1)
            label_entry.insert(0, "")
            self.bind_axis_navigation(label_entry, self.new_y_labels, "y", row, "16x16_vo")
            self.new_y_labels.append(label_entry)

            row_cells = []
            for col in range(self.new_cols):
                entry = tk.Entry(frame_16x16, width=5, font=('Arial', 8))
                entry.grid(row=row + 1, column=col + 1, padx=1, pady=1)
                self.bind_cell_events(entry, self.new_cells, row, col, "16x16_vo")
                row_cells.append(entry)
            self.new_cells.append(row_cells)

        # ----- 16x16 VE TABLE -----
        frame_ve = tk.LabelFrame(bottom_frame, text="Calculated 16x16 Table (VE)")
        frame_ve.pack(side=tk.LEFT, padx=(10, 0), fill=tk.BOTH, expand=True)

        for col in range(self.new_cols):
            entry = tk.Entry(frame_ve, width=5, bg="#d0d0d0", font=('Arial', 8))
            entry.grid(row=0, column=col + 1, padx=1, pady=1)
            entry.insert(0, "")
            self.bind_axis_navigation(entry, self.ve_x_labels, "x", col, "16x16_ve")
            self.ve_x_labels.append(entry)

        for row in range(self.new_rows):
            label_entry = tk.Entry(frame_ve, width=5, bg="#d0d0d0", font=('Arial', 8))
            label_entry.grid(row=row + 1, column=0, padx=1, pady=1)
            label_entry.insert(0, "")
            self.bind_axis_navigation(label_entry, self.ve_y_labels, "y", row, "16x16_ve")
            self.ve_y_labels.append(label_entry)

            row_cells = []
            for col in range(self.new_cols):
                entry = tk.Entry(frame_ve, width=5, font=('Arial', 8))
                entry.grid(row=row + 1, column=col + 1, padx=1, pady=1)
                self.bind_cell_events(entry, self.ve_cells, row, col, "16x16_ve")
                row_cells.append(entry)
            self.ve_cells.append(row_cells)

        # Buttons for VE table
        ve_btn_frame = tk.Frame(frame_ve)
        ve_btn_frame.grid(row=self.new_rows + 2, column=0, columnspan=self.new_cols + 1, pady=5)

        tk.Button(ve_btn_frame, text="Copy VE Table", command=self.copy_ve_table).pack(side=tk.LEFT, padx=5)
        tk.Button(ve_btn_frame, text="Convert VE → VO", command=self.convert_ve_to_vo).pack(side=tk.LEFT, padx=5)

    # ===== UI BINDINGS =====

    def bind_axis_navigation(self, entry, axis_list, axis_type, index, table_name):
        """
        Custom keyboard navigation for axis entries (X/Y labels).
        Works for 8x12, 16x16 VO and 16x16 VE axes.
        """
        def on_key(event):
            total = len(axis_list)

            # TAB cycles through entries in this axis list
            if event.keysym == 'Tab':
                # Shift+Tab goes backwards
                shift_pressed = (event.state & 0x0001) != 0
                if not shift_pressed:
                    next_index = (index + 1) % total
                else:
                    next_index = (index - 1) % total
                axis_list[next_index].focus_set()
                return "break"

            # Horizontal navigation on X axis
            if event.keysym in ('Right', 'Left') and axis_type == "x":
                if event.keysym == 'Right':
                    next_index = min(total - 1, index + 1)
                else:
                    next_index = max(0, index - 1)
                axis_list[next_index].focus_set()
                return "break"

            # Vertical navigation or jump from X axis to first data row
            if event.keysym == 'Down':
                if axis_type == "x":
                    # Jump from RPM axis to first row of data in corresponding table
                    if table_name == "8x12":
                        self.cells[0][index].focus_set()
                    elif table_name == "16x16_vo":
                        self.new_cells[0][index].focus_set()
                    elif table_name == "16x16_ve":
                        self.ve_cells[0][index].focus_set()
                else:
                    # Y axis: move down within axis labels
                    next_index = min(total - 1, index + 1)
                    axis_list[next_index].focus_set()
                return "break"

            if event.keysym == 'Up' and axis_type == "y":
                prev_index = max(0, index - 1)
                axis_list[prev_index].focus_set()
                return "break"

            return None

        entry.bind("<Key>", on_key)

    def bind_cell_events(self, entry, table_list, row, col, table_name):
        """Bind mouse and keyboard events for a data cell."""
        entry.bind("<Button-1>", lambda e, r=row, c=col, t=table_name: self.on_cell_click(e, r, c, t))
        entry.bind("<B1-Motion>", lambda e, r=row, c=col, t=table_name: self.on_cell_drag(e, r, c, t))
        entry.bind("<ButtonRelease-1>", lambda e: self.on_cell_release(e))
        entry.bind("<Key>", lambda e, r=row, c=col, t=table_name: self.on_cell_key(e, r, c, t))

    def bind_events(self):
        """Bind global keyboard shortcuts."""
        self.root.bind("<Control-v>", self.handle_paste)
        self.root.bind("<Control-a>", self.select_all)
        self.root.bind("<Control-c>", self.copy_selection)

    # ===== SELECTION AND TABLE HELPERS =====

    def get_table_from_name(self, table_name):
        if table_name == "8x12":
            return self.cells
        elif table_name == "16x16_vo":
            return self.new_cells
        elif table_name == "16x16_ve":
            return self.ve_cells
        return None

    def get_table_dimensions(self, table_name):
        if table_name == "8x12":
            return self.rows, self.cols
        else:
            return self.new_rows, self.new_cols

    def on_cell_click(self, event, row, col, table_name):
        self.clear_selection()
        self.current_table = table_name
        self.selection_start = (row, col)
        self.update_selection(row, col, row, col)

    def on_cell_drag(self, event, row, col, table_name):
        if self.selection_start is None or self.current_table != table_name:
            return

        table = self.get_table_from_name(table_name)
        rows, cols = self.get_table_dimensions(table_name)

        for r in range(rows):
            for c in range(cols):
                cell = table[r][c]
                x1 = cell.winfo_rootx()
                y1 = cell.winfo_rooty()
                x2 = x1 + cell.winfo_width()
                y2 = y1 + cell.winfo_height()

                if x1 <= event.x_root <= x2 and y1 <= event.y_root <= y2:
                    self.update_selection(self.selection_start[0], self.selection_start[1], r, c)
                    return

    def on_cell_release(self, event):
        pass

    def on_cell_key(self, event, row, col, table_name):
        """
        Keyboard navigation for data cells:
        - Tab / Shift+Tab moves logically through the grid.
        - Arrow keys move up/down/left/right.
        """
        table = self.get_table_from_name(table_name)
        if table is None:
            return None

        rows, cols = self.get_table_dimensions(table_name)

        def clamp(rc, cc):
            rc = max(0, min(rows - 1, rc))
            cc = max(0, min(cols - 1, cc))
            return rc, cc

        if event.keysym == "Tab":
            shift_pressed = (event.state & 0x0001) != 0
            if not shift_pressed:
                new_row, new_col = row, col + 1
                if new_col >= cols:
                    new_col = 0
                    new_row += 1
                    if new_row >= rows:
                        new_row = 0
                new_row, new_col = clamp(new_row, new_col)
            else:
                new_row, new_col = row, col - 1
                if new_col < 0:
                    new_col = cols - 1
                    new_row -= 1
                    if new_row < 0:
                        new_row = rows - 1
                new_row, new_col = clamp(new_row, new_col)

            table[new_row][new_col].focus_set()
            return "break"

        if event.keysym == "Right":
            new_row, new_col = clamp(row, col + 1)
            table[new_row][new_col].focus_set()
            return "break"
        if event.keysym == "Left":
            new_row, new_col = clamp(row, col - 1)
            table[new_row][new_col].focus_set()
            return "break"
        if event.keysym == "Down":
            new_row, new_col = clamp(row + 1, col)
            table[new_row][new_col].focus_set()
            return "break"
        if event.keysym == "Up":
            new_row, new_col = clamp(row - 1, col)
            table[new_row][new_col].focus_set()
            return "break"

        return None

    def update_selection(self, start_row, start_col, end_row, end_col):
        self.clear_selection()

        table = self.get_table_from_name(self.current_table)
        if table is None:
            return

        min_row, max_row = min(start_row, end_row), max(start_row, end_row)
        min_col, max_col = min(start_col, end_col), max(start_col, end_col)

        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                cell = table[r][c]
                cell.configure(bg="#cce5ff")
                self.selected_cells.append((r, c))

    def clear_selection(self):
        if self.current_table:
            table = self.get_table_from_name(self.current_table)
            if table:
                for r, c in self.selected_cells:
                    table[r][c].configure(bg="white")
        self.selected_cells = []

    def select_all(self, event=None):
        """Select all cells in the table where the focused cell lives."""
        focused = self.root.focus_get()

        for table_name, table in [("8x12", self.cells), ("16x16_vo", self.new_cells), ("16x16_ve", self.ve_cells)]:
            rows, cols = self.get_table_dimensions(table_name)
            for r in range(rows):
                for c in range(cols):
                    if table[r][c] == focused:
                        self.current_table = table_name
                        self.selection_start = (0, 0)
                        self.update_selection(0, 0, rows - 1, cols - 1)
                        return "break"
        return None

    def copy_selection(self, event=None):
        """Copy selected cell range to clipboard as tab-separated text."""
        if not self.selected_cells or not self.current_table:
            return

        table = self.get_table_from_name(self.current_table)
        if table is None:
            return

        min_row = min(r for r, c in self.selected_cells)
        max_row = max(r for r, c in self.selected_cells)
        min_col = min(c for r, c in self.selected_cells)
        max_col = max(c for r, c in self.selected_cells)

        lines = []
        for r in range(min_row, max_row + 1):
            row_values = []
            for c in range(min_col, max_col + 1):
                row_values.append(table[r][c].get())
            lines.append('\t'.join(row_values))

        self.root.clipboard_clear()
        self.root.clipboard_append('\n'.join(lines))
        return "break"

    def handle_paste(self, event=None):
        """Paste clipboard into the table where the focused cell is."""
        focused = self.root.focus_get()

        target_table = None
        start_row, start_col = None, None
        rows, cols = 0, 0

        for table_name, table in [("8x12", self.cells),
                                  ("16x16_vo", self.new_cells),
                                  ("16x16_ve", self.ve_cells)]:
            r_count, c_count = self.get_table_dimensions(table_name)
            for r in range(r_count):
                for c in range(c_count):
                    if table[r][c] == focused:
                        target_table = table
                        start_row, start_col = r, c
                        rows, cols = r_count, c_count
                        break
                if target_table is not None:
                    break
            if target_table is not None:
                break

        if target_table is None:
            return

        try:
            clipboard = self.root.clipboard_get()
        except Exception:
            return

        lines = clipboard.strip().split('\n')
        for i, line in enumerate(lines):
            r = start_row + i
            if r >= rows:
                break
            values = line.split('\t')
            for j, value in enumerate(values):
                c = start_col + j
                if c >= cols:
                    break
                cell = target_table[r][c]
                cell.delete(0, tk.END)
                cell.insert(0, value.strip())

        return "break"

    def clear_all(self):
        """Clear all cells in the 8x12 VO table."""
        for row in self.cells:
            for cell in row:
                cell.delete(0, tk.END)

    def copy_ve_table(self):
        """Copy entire 16x16 VE table to clipboard."""
        lines = []
        for row in self.ve_cells:
            values = [cell.get() for cell in row]
            lines.append('\t'.join(values))

        self.root.clipboard_clear()
        self.root.clipboard_append('\n'.join(lines))
        messagebox.showinfo("Copied", "VE table copied to clipboard")

    # ===== AXIS EXTENSION =====

    def extend_x_axis(self, values):
        """Extend RPM axis (X) from 12 columns to 16 using rules for NA/FI."""
        last_rpm = values[-1]
        new_x = []

        if last_rpm == 7000:
            # Classic MS4x-like 320–7000 map extended to 16 cols
            new_x = values[:8].copy()
            current = new_x[-1]
            while len(new_x) < 16:
                current += 500
                if current > 7000:
                    current = 7000
                new_x.append(current)
                if current == 7000:
                    break
            while len(new_x) < 16:
                new_x.append(7000)

        elif last_rpm <= 7500:
            # Slightly higher rev limit – extend with 500 rpm steps
            new_x = values[:3].copy()
            current = new_x[-1]
            while len(new_x) < 16:
                current += 500
                if current > last_rpm:
                    current = last_rpm
                new_x.append(current)
                if current == last_rpm:
                    break
            while len(new_x) < 16:
                new_x.append(last_rpm)

        else:
            # High-rev setup – custom final steps up to 8000
            new_x = values[:3].copy()
            current = new_x[-1]
            current += 700
            new_x.append(current)
            current += 600
            new_x.append(current)

            while len(new_x) < 16:
                current += 500
                if current > 8000:
                    current = 8000
                new_x.append(current)
                if current == 8000:
                    break
            while len(new_x) < 16:
                new_x.append(8000)

        return new_x

    def extend_y_axis(self, values):
        """Extend MAP axis (Y) from 8 rows to 16 rows, NA vs FI preset."""
        max_kpa = max(values)
        if max_kpa <= 125:
            # NA preset
            return [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 60, 70, 80, 90, 105, 125]
        else:
            # Boost preset
            return [10, 20, 30, 35, 40, 45, 50, 60, 70, 80, 90, 105, 125, 200, 250, 300]

    # ===== INTERPOLATION + VERTICAL SMOOTHING =====

    def bilinear_interpolate_point(self, x, y, x_vals, y_vals, grid):
        """
        Bilinear interpolation of VO value at point (x, y)
        from original 8x12 grid defined by x_vals, y_vals and grid[row][col].
        """
        cols = len(x_vals)
        rows = len(y_vals)

        # Find bracketing x indices
        if x <= x_vals[0]:
            k0 = k1 = 0
        elif x >= x_vals[-1]:
            k0 = k1 = cols - 1
        else:
            k0 = 0
            for k in range(cols - 1):
                if x_vals[k] <= x <= x_vals[k + 1]:
                    k0 = k
                    break
            k1 = k0 + 1

        # Find bracketing y indices
        if y <= y_vals[0]:
            l0 = l1 = 0
        elif y >= y_vals[-1]:
            l0 = l1 = rows - 1
        else:
            l0 = 0
            for l in range(rows - 1):
                if y_vals[l] <= y <= y_vals[l + 1]:
                    l0 = l
                    break
            l1 = l0 + 1

        x0, x1 = x_vals[k0], x_vals[k1]
        y0, y1 = y_vals[l0], y_vals[l1]

        q00 = grid[l0][k0]
        q10 = grid[l0][k1]
        q01 = grid[l1][k0]
        q11 = grid[l1][k1]

        # Edge cases where interpolation reduces to 1D or a single point
        if k0 == k1 and l0 == l1:
            return q00
        if k0 == k1:
            if y1 == y0:
                return q00
            t = (y - y0) / (y1 - y0)
            return q00 + (q01 - q00) * t
        if l0 == l1:
            if x1 == x0:
                return q00
            t = (x - x0) / (x1 - x0)
            return q00 + (q10 - q00) * t

        # Full bilinear interpolation
        tx = 0 if x1 == x0 else (x - x0) / (x1 - x0)
        ty = 0 if y1 == y0 else (y - y0) / (y1 - y0)

        a = q00 * (1 - tx) + q10 * tx
        b = q01 * (1 - tx) + q11 * tx
        return a * (1 - ty) + b * ty

    def vertical_smooth(self, grid, fixed_mask, beta=0.6, iterations=2):
        """
        Stronger vertical smoothing along MAP (Y) axis.

        For each non-fixed cell, pull it towards the average of the cell above
        and below. beta controls how strong the smoothing is (0..1), iterations
        how many passes we do.
        """
        rows = len(grid)
        cols = len(grid[0]) if rows > 0 else 0

        for _ in range(iterations):
            new_grid = [row[:] for row in grid]
            for r in range(1, rows - 1):  # skip first/last row (need both neighbors)
                for c in range(cols):
                    if fixed_mask[r][c]:
                        continue

                    up = grid[r - 1][c]
                    down = grid[r + 1][c]
                    avg = (up + down) / 2.0

                    new_grid[r][c] = grid[r][c] * (1.0 - beta) + avg * beta

            grid = new_grid

        return grid

    def fill_data_cells(self, old_x, old_y, new_x, new_y):
        """
        Create 16x16 VO grid:
        1) Bilinear interpolation from original 8x12
        2) Vertical smoothing along MAP axis while keeping original 8x12 points fixed
        """
        # Clear 16x16 VO table
        for row in self.new_cells:
            for cell in row:
                cell.delete(0, tk.END)

        # Read original 8x12 VO grid
        try:
            old_grid = [
                [float(self.cells[r][c].get()) for c in range(self.cols)]
                for r in range(self.rows)
            ]
        except ValueError:
            messagebox.showerror("Error", "All 8x12 VO cells must contain valid numbers")
            return None

        # 1) Initial 16x16 grid from pure bilinear interpolation
        data_grid = [[None for _ in range(self.new_cols)] for _ in range(self.new_rows)]

        for r in range(self.new_rows):
            y = new_y[r]
            for c in range(self.new_cols):
                x = new_x[c]
                val = self.bilinear_interpolate_point(x, y, old_x, old_y, old_grid)
                data_grid[r][c] = val

        # 2) Build mask of original 8x12 points in 16x16 grid
        fixed_mask = [[False for _ in range(self.new_cols)] for _ in range(self.new_rows)]

        for r_old, y_val in enumerate(old_y):
            for c_old, x_val in enumerate(old_x):
                try:
                    r_new = new_y.index(y_val)
                    c_new = new_x.index(x_val)
                except ValueError:
                    # Original axis value not present in extended axis
                    continue

                fixed_mask[r_new][c_new] = True
                data_grid[r_new][c_new] = old_grid[r_old][c_old]

        # 3) Stronger vertical smoothing to reduce large jumps between MAP rows
        data_grid = self.vertical_smooth(data_grid, fixed_mask, beta=0.6, iterations=2)

        # 4) Write back to GUI 16x16 VO table
        for row in range(self.new_rows):
            for col in range(self.new_cols):
                val = data_grid[row][col]
                if val is not None:
                    self.new_cells[row][col].insert(0, f"{val:.3f}")

        return data_grid

    # ===== MAIN OPERATIONS =====

    def generate_16x16(self):
        """Generate extended 16x16 VO table from 8x12 input."""
        try:
            x_values = [float(entry.get()) for entry in self.x_labels]
        except ValueError:
            messagebox.showerror("Error", "X-axis must contain numbers")
            return

        try:
            y_values = [float(entry.get()) for entry in self.y_labels]
        except ValueError:
            messagebox.showerror("Error", "Y-axis must contain numbers")
            return

        self.new_x = self.extend_x_axis(x_values)
        self.new_y = self.extend_y_axis(y_values)

        max_kpa = max(y_values)
        mode = "NA" if max_kpa <= 125 else "Forced Induction"

        for i, entry in enumerate(self.new_x_labels):
            entry.delete(0, tk.END)
            entry.insert(0, f"{self.new_x[i]:.0f}")

        for i, entry in enumerate(self.new_y_labels):
            entry.delete(0, tk.END)
            entry.insert(0, f"{self.new_y[i]:.0f}")

        self.vo_grid = self.fill_data_cells(x_values, y_values, self.new_x, self.new_y)
        if self.vo_grid is not None:
            messagebox.showinfo("Done", f"16x16 VO table generated\nMode: {mode}")

    def calculate_ve(self):
        """Calculate 16x16 VE table from 16x16 VO grid using thermodynamic formula."""
        if self.vo_grid is None or self.new_x is None or self.new_y is None:
            messagebox.showerror("Error", "Generate 16x16 table first")
            return

        try:
            displacement = float(self.displacement_var.get())
        except ValueError:
            messagebox.showerror("Error", "Displacement must be a number")
            return

        try:
            iat = float(self.iat_var.get())
        except ValueError:
            messagebox.showerror("Error", "IAT must be a number")
            return

        # Clear VE cells
        for row in self.ve_cells:
            for cell in row:
                cell.delete(0, tk.END)

        # Copy axes to VE table
        for i, entry in enumerate(self.ve_x_labels):
            entry.delete(0, tk.END)
            entry.insert(0, f"{self.new_x[i]:.0f}")

        for i, entry in enumerate(self.ve_y_labels):
            entry.delete(0, tk.END)
            entry.insert(0, f"{self.new_y[i]:.0f}")

        # VE calculation
        for row in range(self.new_rows):
            for col in range(self.new_cols):
                vo = self.vo_grid[row][col]
                if vo is None:
                    continue

                rpm = self.new_x[col]
                map_kpa = self.new_y[row]

                if map_kpa == 0:
                    continue

                numerator = vo * rpm * 8.314 * (iat + 273.15) * 120
                denominator = 5555 * map_kpa * displacement * rpm * 28.9 * 3.6
                ve = numerator / denominator

                self.ve_cells[row][col].insert(0, f"{ve:.3f}")

        messagebox.showinfo("Done", "VE table calculated")

    def convert_ve_to_vo(self):
        """
        Convert 16x16 VE table back to VO using the inverse of the VE formula.
        Result is written into 16x16 VO table and stored in self.vo_grid.
        """
        if self.new_x is None or self.new_y is None:
            # Try to rebuild axes from VE header if user pasted VE only
            try:
                self.new_x = [float(e.get()) for e in self.ve_x_labels]
                self.new_y = [float(e.get()) for e in self.ve_y_labels]
            except ValueError:
                messagebox.showerror(
                    "Error",
                    "Invalid VE axes. Generate 16x16 table first or fill VE axes."
                )
                return

        try:
            displacement = float(self.displacement_var.get())
        except ValueError:
            messagebox.showerror("Error", "Displacement must be a number")
            return

        try:
            iat = float(self.iat_var.get())
        except ValueError:
            messagebox.showerror("Error", "IAT must be a number")
            return

        # Constants from the VE formula
        const_num = 5555 * displacement * 28.9 * 3.6
        const_den = 8.314 * (iat + 273.15) * 120

        # Clear VO 16x16 table
        for row in self.new_cells:
            for cell in row:
                cell.delete(0, tk.END)

        # Sync VO axes
        for i, entry in enumerate(self.new_x_labels):
            entry.delete(0, tk.END)
            entry.insert(0, f"{self.new_x[i]:.0f}")
        for i, entry in enumerate(self.new_y_labels):
            entry.delete(0, tk.END)
            entry.insert(0, f"{self.new_y[i]:.0f}")

        self.vo_grid = [[None for _ in range(self.new_cols)] for _ in range(self.new_rows)]

        for r in range(self.new_rows):
            map_kpa = self.new_y[r]
            if map_kpa == 0:
                continue

            for c in range(self.new_cols):
                ve_str = self.ve_cells[r][c].get().strip()
                if not ve_str:
                    continue

                try:
                    ve_val = float(ve_str)
                except ValueError:
                    continue

                vo = ve_val * const_num * map_kpa / const_den

                self.vo_grid[r][c] = vo
                self.new_cells[r][c].insert(0, f"{vo:.3f}")

        messagebox.showinfo("Done", "VO table calculated from VE")


if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = TableEditor(root)
        root.mainloop()
    except Exception as e:
        print(f"Error: {e}")
        input("Press Enter to close...")
