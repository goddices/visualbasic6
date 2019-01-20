Option Explicit

'ScaleMode.
Global Const PIXELS = 3

'Show parameter for form.
Global Const MODAL = 1

'BorderStyle for form and label.
Global Const NONE = 0
Global Const FIXED_SINGLE = 1

'Alignment for label.
Global Const CENTER = 2

'Fixed window size in pixels (solitaire default on 800x600).
Global Const XWIN_SIZE = 593
Global Const YWIN_SIZE = 429

'State constants.
Global Const WAITING = 0
Global Const DRAGGING = 1

'Card dimensions.
Global Const CARD_SIZEX = 71
Global Const CARD_SIZEY = 96

'Coordinates of source rectangle at top left.
Global Const SOURCE_LEFT = 11
Global Const SOURCE_TOP = 5

'Coordinates of first home rectangle.
Global Const HOME_LEFT = 257
Global Const HOME_TOP = 5

'Offset between home rectangles.
Global Const HOME_OFFSET = 82

'Coordinates of first holding rectangle in second row.
Global Const HOLD_LEFT = 11
Global Const HOLD_TOP = 107

'Offset between holding rectangles.
Global Const HOLD_OFFSET = 82

'Number of pixels skipped in card's auto-return trip.
'The larger SKIP_FACTOR is, the faster.
Global Const SKIP_FACTOR = 60

'BOUNDARY is a calculated value used in HomeHoldOverlap() to determine
'which of two HomeRects is overlapped the most. Although hard-coded
'here, BOUNDARY depends entirely on CARD_SIZEX and HOME_OFFSET:
'  GAP = HOME_OFFSET - CARD_SIZEX = 82 - 71 = 11
'  MID_OLAP = (CARD_SIZEX - GAP) / 2 = (71 - 11) / 2 = 60/2 = 30
'  BOUNDARY = CARD_SIZEX - MID_OLAP = 71 - 30 = 41
Global Const BOUNDARY = 41

'Height of info strip at bottom of form (in pixels).
Global Const INFO_LINE_HEIGHT = 18

'Offset of top of info strip from bottom of form (in pixels).
'This constant is essentially the height of form's caption bar +
'height of form's menu bar + INFO_LINE_HEIGHT
Global Const INFO_LINE_OFFSET = 57

'Key variable telling state of drag (WAITING or DRAGGING).
Global State%

'Button down point from top left of card when dragging.
Global x0%, y0%

