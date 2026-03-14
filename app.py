from flask import Flask
from config import Config
from routes.main import main_bp
from routes.isocal import isocal_bp


def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    app.register_blueprint(main_bp)
    app.register_blueprint(isocal_bp)

    @app.errorhandler(413)
    def too_large(e):
        from flask import render_template, flash
        flash("File too large. Maximum upload size is 10 MB.", "error")
        from forms.isocal_form import IsocalForm
        return render_template("tools/isocal.html", form=IsocalForm()), 413

    return app


if __name__ == "__main__":
    app = create_app()
    app.run(debug=True)
