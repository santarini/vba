'Move to the edge of the current data region from selection
    'To the right
        Selection.End(xlToRight).Select
    'To the left
        Selection.End(xlToLeft).Select
    'To the down
        Selection.End(xlDown).Select
    'To the up
        Selection.End(xlUp).Select


'Select until the edge fo the current data region
    'To the right
        Range(Selection, Selection.End(xlToRight)).Select
    'To the left
        Range(Selection, Selection.End(xlToLeft)).Select
    'To the down
        Range(Selection, Selection.End(xlDown)).Select
    'To the up
        Range(Selection, Selection.End(xlUp)).Select
