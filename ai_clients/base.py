from abc import ABC, abstractmethod

class AIClient(ABC):
    @abstractmethod
    def process(self, prompt: str) -> str:
        pass
