from dataclasses import dataclass
from typing import List, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPlainTextEdit,
    QRadioButton,
    QCheckBox,
    QVBoxLayout,
)


@dataclass
class NameResolution:
    names: List[str]
    cache_original: bool
    add_new: bool = False
    persist_alias: bool = False
    merge_with_prev: bool = False
    merge_with_next: bool = False


class NameResolutionDialog(QDialog):
    def __init__(
        self,
        name: str,
        suggestions: List[str],
        line_text: str = "",
        prev_token: str = "",
        next_token: str = "",
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Resolve Name")
        self._result: Optional[NameResolution] = None
        self._merge_prev_value: Optional[str] = None
        self._merge_next_value: Optional[str] = None

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"Unknown name: {name}"))

        if line_text:
            layout.addWidget(QLabel("Line"))
            line_view = QPlainTextEdit()
            line_view.setReadOnly(True)
            line_view.setPlainText(line_text)
            line_view.setMinimumHeight(80)
            layout.addWidget(line_view)

        self.button_group = QButtonGroup(self)

        if prev_token:
            merged = f"{prev_token}{name}"
            self._merge_prev_value = merged
            self.merge_prev_button = QRadioButton(
                f"Merge with previous token: {merged}"
            )
            self.button_group.addButton(self.merge_prev_button)
            layout.addWidget(self.merge_prev_button)
        else:
            self.merge_prev_button = None

        if next_token:
            merged = f"{name}{next_token}"
            self._merge_next_value = merged
            self.merge_next_button = QRadioButton(
                f"Merge with next token: {merged}"
            )
            self.button_group.addButton(self.merge_next_button)
            layout.addWidget(self.merge_next_button)
        else:
            self.merge_next_button = None

        for idx, suggestion in enumerate(suggestions, start=1):
            btn = QRadioButton(f"Use suggestion {idx}: {suggestion}")
            self.button_group.addButton(btn)
            layout.addWidget(btn)

        self.custom_button = QRadioButton("Enter a different name")
        self.button_group.addButton(self.custom_button)
        layout.addWidget(self.custom_button)
        self.custom_input = QLineEdit()
        self.custom_input.setPlaceholderText("Type a name")
        layout.addWidget(self.custom_input)

        self.split_button = QRadioButton("Split into two names")
        self.button_group.addButton(self.split_button)
        layout.addWidget(self.split_button)
        split_row = QHBoxLayout()
        self.split_first = QLineEdit()
        self.split_first.setPlaceholderText("First name")
        self.split_second = QLineEdit()
        self.split_second.setPlaceholderText("Second name")
        split_row.addWidget(self.split_first)
        split_row.addWidget(self.split_second)
        layout.addLayout(split_row)

        self.add_new_button = QRadioButton("Add as a new name")
        self.button_group.addButton(self.add_new_button)
        layout.addWidget(self.add_new_button)
        self.add_new_input = QLineEdit()
        self.add_new_input.setPlaceholderText("New name")
        layout.addWidget(self.add_new_input)

        self.discard_button = QRadioButton("Discard")
        self.button_group.addButton(self.discard_button)
        layout.addWidget(self.discard_button)

        self.persist_alias_checkbox = QCheckBox("Save this mapping to name_aliases.json")
        self.persist_alias_checkbox.setChecked(True)
        layout.addWidget(self.persist_alias_checkbox)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        if self.button_group.buttons():
            self.button_group.buttons()[0].setChecked(True)

    def _on_accept(self) -> None:
        selected = self.button_group.checkedButton()
        if selected:
            text = selected.text()
            if self.merge_prev_button and selected == self.merge_prev_button:
                if self._merge_prev_value:
                    self._result = NameResolution(
                        [self._merge_prev_value],
                        cache_original=False,
                        persist_alias=False,
                        merge_with_prev=True,
                    )
                    self.accept()
                    return
            if self.merge_next_button and selected == self.merge_next_button:
                if self._merge_next_value:
                    self._result = NameResolution(
                        [self._merge_next_value],
                        cache_original=False,
                        persist_alias=False,
                        merge_with_next=True,
                    )
                    self.accept()
                    return
            if text.startswith("Use suggestion"):
                suggestion = text.split(":", 1)[1].strip()
                self._result = NameResolution(
                    [suggestion],
                    cache_original=True,
                    persist_alias=self.persist_alias_checkbox.isChecked(),
                )
                self.accept()
                return

        if self.custom_button.isChecked():
            value = self.custom_input.text().strip()
            if value:
                self._result = NameResolution(
                    [value],
                    cache_original=True,
                    persist_alias=self.persist_alias_checkbox.isChecked(),
                )
                self.accept()
                return

        if self.split_button.isChecked():
            first = self.split_first.text().strip()
            second = self.split_second.text().strip()
            if first and second:
                self._result = NameResolution(
                    [first, second],
                    cache_original=False,
                    persist_alias=False,
                )
                self.accept()
                return

        if self.add_new_button.isChecked():
            value = self.add_new_input.text().strip()
            if value:
                self._result = NameResolution(
                    [value],
                    cache_original=True,
                    add_new=True,
                    persist_alias=self.persist_alias_checkbox.isChecked(),
                )
                self.accept()
                return

        if self.discard_button.isChecked():
            self._result = None
            self.accept()
            return

        self.reject()

    def get_resolution(self) -> Optional[NameResolution]:
        return self._result
