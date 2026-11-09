"""
local_app — local browser front end for the RadioBefundungszahlConverter.

Lives only on the branch `local` and is never merged into `main`. It adds files
and edits none: everything it needs from the converter comes from `core/`.
"""

# Shown in the footer of the page, so that a reported error code can be tied to a build.
__version__ = "0.1.0"
