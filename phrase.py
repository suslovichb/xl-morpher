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

    def morph(self, case, number_category):
        analyzer = pymorphy2.MorphAnalyzer()
        morphed_phrase_items = []
        for item in self.phrase_splitted:
            parsed_item = analyzer.parse(item)[0]
            if parsed_item.tag.POS == 'PREP':
                morphed_phrase_items +=  self.phrase_splitted[self.phrase_splitted.index(item)::]
                break
            if parsed_item.tag.POS in self.MORPHABLE_GRAMMEMES:
                morphed_item = parsed_item.inflect({number_category, case})
                # morphed_phrase_items.append(parsed_item.inflect({number_category, case}).word)
                if morphed_item:
                    morphed_phrase_items.append(morphed_item.word)
                else:
                    morphed_phrase_items.append(item)
            else:
                morphed_phrase_items.append(item)
        return ''.join(morphed_phrase_items)
