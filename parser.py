import re

class MCQParser:
    def __init__(self):
        pass
        
    def parse(self, raw_text: str) -> dict:
        """
        Parses raw OCR text into Options (A, B, C, D).
        Very robust regex to find A, B, C, D anywhere.
        """
        result = {
            "Option A": "", "Option B": "", "Option C": "", "Option D": "",
            "Question": raw_text.strip() # Fallback
        }
        
        # Clean text: replace multiple spaces with single space
        clean_text = re.sub(r'\s+', ' ', raw_text)
        
        # Pattern: Look for (A) or A) or A. or A: (Case Insensitive)
        # We use a lookahead to find the next option or end of string
        pattern = re.compile(r'[\(\[]?\s*([A-Da-d])[\s\.\)\:\]]+')
        
        matches = list(pattern.finditer(clean_text))
        
        if not matches:
            return result
            
        for i in range(len(matches)):
            opt_letter = matches[i].group(1).upper()
            start = matches[i].end()
            end = matches[i+1].start() if i + 1 < len(matches) else len(clean_text)
            
            opt_text = clean_text[start:end].strip()
            col_name = f"Option {opt_letter}"
            if col_name in result:
                result[col_name] = opt_text
                
        return result
        
        result["Option A"] = " ".join(options["A"])
        result["Option B"] = " ".join(options["B"])
        result["Option C"] = " ".join(options["C"])
        result["Option D"] = " ".join(options["D"])
        
        return result

if __name__ == "__main__":
    test_text = '''1. What is the capital of France?
A) Paris
B) London
C) Rome
D) Berlin'''
    parser = MCQParser()
    print(parser.parse(test_text))
