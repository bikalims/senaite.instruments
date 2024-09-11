from senaite.core.browser.form.adapters.data_import import EditForm as EF
import os


class EditForm(EF):

    def get_default_import_template(self):
        """Returns the path of the default import template
        """
        import senaite.instruments.instruments
        path = os.path.dirname(senaite.instruments.instruments.__file__)
        template = "instrument.pt"
        return os.path.join(path, template)
