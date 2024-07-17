print("Hello, Convertor!")
try:
    import docx2pdf as dp

    print("Success")
    input()
except ImportError as I:
    print(I)
    input()
    print("Failure")
finally:
    print("DONE")
    input()
