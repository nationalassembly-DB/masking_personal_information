"""
파일을 masking 합니다.
"""


def _masking_hwp(hwp, text):
    """hwp문서를 마스킹합니다"""
    text_num = len(text.replace('\r', '').replace('\n', ''))

    hwp.Run("Select")
    hwp.Run("Select")
    hwp.HAction.GetDefault(
        "InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = '*' * text_num
    hwp.HAction.Execute(
        "InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.Run("Cancel")
