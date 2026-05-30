"""
FormatX 启动页 — 暖白遮罩 + 品牌文字 + 淡出动画
"""
import customtkinter as ctk

_COLOR_BG  = "#F9F7F2"
_COLOR_TEXT = "#37352F"
_COLOR_BAR = "#F5F3ED"


class SplashScreen(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()
        self.overrideredirect(True)
        self.configure(fg_color=_COLOR_BG)

        w, h = 400, 320
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw - w)//2}+{(sh - h)//2}")
        self.attributes("-alpha", 0.95)

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        center = ctk.CTkFrame(self, fg_color="transparent")
        center.grid(row=0, column=0)

        ctk.CTkLabel(
            center, text="FormatX",
            font=ctk.CTkFont(family="Microsoft YaHei", size=36, weight="normal"),
            text_color=_COLOR_TEXT
        ).pack(pady=(30, 5))

        ctk.CTkLabel(
            center,
            text="你是我荒芜岁月里，\n突然腾起的群鸦。",
            font=ctk.CTkFont(family="Microsoft YaHei", size=14, slant="italic"),
            text_color="#9B8D82", justify="center"
        ).pack(pady=(0, 18))

        # 分割线
        ctk.CTkFrame(center, height=2, fg_color=_COLOR_BAR, width=200).pack(pady=(0, 20))

        # 加载行：文字 + ✻ 并排
        load_row = ctk.CTkFrame(center, fg_color="transparent")
        load_row.pack()

        ctk.CTkLabel(
            load_row, text="正在启动排版引擎...",
            font=ctk.CTkFont(family="Microsoft YaHei", size=12),
            text_color="#B0A8A0"
        ).pack(side="left")

        self._spinner = ctk.CTkLabel(
            load_row, text=" ✻",
            font=ctk.CTkFont(family="Microsoft YaHei", size=14),
            text_color="#B0A8A0"
        )
        self._spinner.pack(side="left")

        self.lift()
        self.focus_force()
        self._pulse_text()

    def _pulse_text(self):
        """呼吸式脉动 + ✻ 旋转感"""
        import math
        self._pulse_phase = getattr(self, '_pulse_phase', 0) + 1
        t = self._pulse_phase
        alpha_val = 0.85 + 0.15 * (1 + math.sin(t * 0.15)) / 2
        self.attributes("-alpha", alpha_val)
        # ✻ 循环切换 Unicode 旋转字符
        spin_chars = ['✻', '⟳', '✧', '✦', '✻', '◌', '✧', '✦']
        char = spin_chars[self._pulse_phase % len(spin_chars)]
        self._spinner.configure(text=f" {char}")
        self._pulse_id = self.after(120, self._pulse_text)

    def fade_out(self):
        if hasattr(self, '_pulse_id'):
            self.after_cancel(self._pulse_id)
        steps = 20
        for i in range(steps + 1):
            t = i / steps
            # ease-in-out: 慢→快→慢
            if t < 0.5:
                eased = 2 * t * t
            else:
                eased = -1 + (4 - 2 * t) * t
            self.attributes("-alpha", 0.95 * (1 - eased))
            self.update()
            self.after(15)
        self.destroy()
