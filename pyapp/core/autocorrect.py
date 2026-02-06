from typing import Dict, Iterable, List, Optional

import textdistance


class Autocorrecter:
    def __init__(self, words: Iterable[str]) -> None:
        self.vocab: Dict[str, str] = {}
        for word in words:
            self.add_word(word)

    def add_word(self, word: str, display: Optional[str] = None) -> None:
        if not word:
            return
        key = word.lower()
        if key in self.vocab:
            return
        self.vocab[key] = display or word

    def correct(self, input_word: str) -> List[str]:
        input_word = input_word.lower()
        similarities = [
            (v, textdistance.cosine(v, input_word)) for v in self.vocab
        ]
        similarities.sort(key=lambda item: item[1])
        return [self.vocab[v] for v, _ in reversed(similarities)][:5]
