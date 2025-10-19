import sys
from docx2pdf import convert

def main():
    if len(sys.argv) != 3:
        print("Usage: convert.py input_file output_file")
        sys.exit(1)
        
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    try:
        convert(input_file, output_file)
        sys.exit(0)
    except Exception as e:
        print(f"Error converting file: {str(e)}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()