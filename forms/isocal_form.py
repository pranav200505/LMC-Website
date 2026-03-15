from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileAllowed, FileRequired
from wtforms import FloatField, RadioField, StringField
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

    # Chart customization mode
    chart_mode    = RadioField(
                        "Chart Mode",
                        choices=[("auto", "Automatic"), ("custom", "Customize")],
                        default="auto")

    # ── Heat Flow chart customization ──
    hf_chart_title   = StringField("Chart Title",       validators=[Optional()], default="")
    hf_x_title       = StringField("X-Axis Title",      validators=[Optional()], default="Time (Days)")
    hf_y_title       = StringField("Y-Axis Title",      validators=[Optional()], default="Heat Flow (mW/g)")
    hf_line_thickness = FloatField("Line Thickness (pt)", validators=[Optional(), NumberRange(min=0.5, max=10)], default=2.25)
    hf_line_color    = StringField("Line Color",         validators=[Optional()], default="#ED7D31")
    hf_x_font_size   = FloatField("X-Axis Font Size (pt)", validators=[Optional(), NumberRange(min=6, max=24)], default=10.0)
    hf_x_font_color  = StringField("X-Axis Font Color", validators=[Optional()], default="#000000")
    hf_y_font_size   = FloatField("Y-Axis Font Size (pt)", validators=[Optional(), NumberRange(min=6, max=24)], default=10.0)
    hf_y_font_color  = StringField("Y-Axis Font Color", validators=[Optional()], default="#000000")
    hf_border_color  = StringField("Border Color",      validators=[Optional()], default="#FFFFFF")

    # ── Heat chart customization ──
    h_chart_title    = StringField("Chart Title",       validators=[Optional()], default="")
    h_x_title        = StringField("X-Axis Title",      validators=[Optional()], default="Time (Days)")
    h_y_title        = StringField("Y-Axis Title",      validators=[Optional()], default="Heat (J/g)")
    h_line_thickness = FloatField("Line Thickness (pt)", validators=[Optional(), NumberRange(min=0.5, max=10)], default=2.25)
    h_line_color     = StringField("Line Color",         validators=[Optional()], default="#4472C4")
    h_x_font_size    = FloatField("X-Axis Font Size (pt)", validators=[Optional(), NumberRange(min=6, max=24)], default=10.0)
    h_x_font_color   = StringField("X-Axis Font Color", validators=[Optional()], default="#000000")
    h_y_font_size    = FloatField("Y-Axis Font Size (pt)", validators=[Optional(), NumberRange(min=6, max=24)], default=10.0)
    h_y_font_color   = StringField("Y-Axis Font Color", validators=[Optional()], default="#000000")
    h_border_color   = StringField("Border Color",      validators=[Optional()], default="#FFFFFF")
