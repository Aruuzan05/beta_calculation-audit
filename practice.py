import pandas as pd
import akshare as ak


# def format_hk_ticker(ticker: str) -> str:
#     ticker = str(ticker).strip()
#     if "." in ticker:
#         ticker = ticker.split(".")[0]
#     return ticker.zfill(5)

# def format_a_ticker(ticker: str) -> str:
#     ticker = str(ticker).strip()
#     if ticker.startswith("6"):
#         return f"sh{ticker}"
#     elif ticker.startswith("0") or ticker.startswith("3"):
#         return f"sz{ticker}"
#     elif ticker.startswith("4") or ticker.startswith("8"):
#         return f"bj{ticker}" if len(ticker) == 6 else f"sz{ticker}"
#     return ticker

