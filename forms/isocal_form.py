from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileAllowed, FileRequired
from wtforms import FloatField, RadioField
from wtforms.validators import DataRequired, Optional, NumberRange


class IsocalForm(FlaskForm):
    raw_file      = FileField(
                        "Raw .xlsx File",
                        validators=[FileRequired(), FileAllowed(["xlsx"], "Only .xlsx files are accepted.")])

    # Required mix design parameters
    paste         = FloatField("Paste (g)",    validators=[DataRequired(), NumberRange(min=0.001, message="Must be > 0")])
    clinker       = FloatField("Clinker (g)",  validators=[DataRequired(), NumberRange(min=0.001, message="Must be > 0")])
    water         = FloatField("Water (g)",    validators=[DataRequired(), NumberRange(min=0.001, message="Must be > 0")])

    # Optional components (default 0)
    gypsum        = FloatField("Gypsum (g)",    validators=[Optional(), NumberRange(min=0)], default=0.0)
    ns            = FloatField("NS (g)",         validators=[Optional(), NumberRange(min=0)], default=0.0)
    noh           = FloatField("NOH (g)",        validators=[Optional(), NumberRange(min=0)], default=0.0)
    limestone     = FloatField("Limestone (g)",  validators=[Optional(), NumberRange(min=0)], default=0.0)
    cc            = FloatField("CC (g)",         validators=[Optional(), NumberRange(min=0)], default=0.0)

    # Cutoff time
    cutoff_mode   = RadioField(
                        "Cutoff Mode",
                        choices=[("auto", "Auto (120 min)"), ("manual", "Manual")],
                        default="auto")
    cutoff_value  = FloatField(
                        "Cutoff (min)",
                        validators=[Optional(), NumberRange(min=1, message="Cutoff must be at least 1 min")])
