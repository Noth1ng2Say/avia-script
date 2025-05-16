import os
import Unprofitability

class Main:
    def __init__(self):
        folder_path = 'C:/Users/m_krylov/Desktop'
        try:
            up = Unprofitability.Unprofitability(folder_path)
            print("Обработка данных завершена успешно!")
        except Exception as e:
            print(f"Произошла ошибка: {str(e)}")

if __name__ == "__main__":
    main_app = Main()




