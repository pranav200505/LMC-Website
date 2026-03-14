import os
import shutil
import tempfile

from flask import Blueprint, render_template, send_file, flash, after_this_request
from werkzeug.utils import secure_filename

from forms.isocal_form import IsocalForm
from processing.isocal import run_isocal

isocal_bp = Blueprint("isocal", __name__)


@isocal_bp.route("/tools/isocal", methods=["GET", "POST"])
def isocal():
    form = IsocalForm()

    if form.validate_on_submit():
        tmp_dir = tempfile.mkdtemp(prefix="lmc_isocal_")
        try:
            upload = form.raw_file.data
            original_name = secure_filename(upload.filename)
            input_path = os.path.join(tmp_dir, original_name)
            upload.save(input_path)

            stem = os.path.splitext(original_name)[0]
            output_name = f"{stem}_analysis.xlsx"
            output_path = os.path.join(tmp_dir, output_name)

            cutoff = None if form.cutoff_mode.data == "auto" else form.cutoff_value.data

            run_isocal(
                input_path=input_path,
                output_path=output_path,
                original_filename=original_name,
                paste=form.paste.data,
                clinker=form.clinker.data,
                gypsum=form.gypsum.data or 0.0,
                ns=form.ns.data or 0.0,
                noh=form.noh.data or 0.0,
                limestone=form.limestone.data or 0.0,
                cc=form.cc.data or 0.0,
                water=form.water.data,
                cutoff=cutoff,
            )

        except Exception as e:
            shutil.rmtree(tmp_dir, ignore_errors=True)
            flash(str(e), "error")
            return render_template("tools/isocal.html", form=form)

        @after_this_request
        def cleanup(response):
            shutil.rmtree(tmp_dir, ignore_errors=True)
            return response

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    return render_template("tools/isocal.html", form=form)
