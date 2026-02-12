import logging
from src.ui import InvoiceApp

def main():
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    app = InvoiceApp()
    app.run()

if __name__ == "__main__":
    main()
