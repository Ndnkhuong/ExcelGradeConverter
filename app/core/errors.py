"""Lỗi được viết để hiển thị trực tiếp cho người dùng."""


class AppError(Exception):
    def __init__(self, message: str):
        super().__init__(message)
        self.message = message
