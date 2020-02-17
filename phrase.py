import pymorphy2

class Phrase:
    WORD_DELIMITERS = [',', ';', ':', ' ']
    MORPHABLE_GRAMMEMES = ['NOUN', 'ADJF', 'ADJS', 'COMP', 'VERB', 'INFN', 'PRTF', 'PRTS', 'NUMR', 'NPRO']

    def __init__(self, phrase_raw):
        self.phrase_raw = phrase_raw
        self.phrase_splitted = self.split_phrase(phrase_raw)

    def split_phrase(self, phrase):
        for delimiter in self.WORD_DELIMITERS:
            phrase = phrase.replace(delimiter, '|' + delimiter + '|')
        phrase = phrase.replace('||', '|')
        return phrase.split('|')

    def morph(self, number_category, case):
        analyzer = pymorphy2.MorphAnalyzer()
        morphed_phrase_items = []
        for item in self.phrase_splitted:
            parsed_item = analyzer.parse(item)[0]
            if parsed_item.tag.POS in self.MORPHABLE_GRAMMEMES:
                morphed_phrase_items.append(parsed_item.inflect({number_category, case}).word)
            else:
                morphed_phrase_items.append(item)
        return morphed_phrase_items
