import StringIO
import os
import xlrd

from django import forms
from django.utils.translation import ugettext_lazy as _

IMPORT_FILE_TYPES = ['.xls', ]

class XlsInputForm(forms.Form):
    input_excel = forms.FileField(required=True,
        label=_("Upload the Excel file to import to the system."))

    def clean(self):
        data = super(XlsInputForm, self).clean()

        if 'input_excel' not in data:
            raise forms.ValidationError(_('The Excel file is required to proceed.'))

        input_excel = data['input_excel']
        extension = os.path.splitext( input_excel.name )[1]
        if not (extension in IMPORT_FILE_TYPES):
            raise forms.ValidationError(
                _(u'%s is not a valid Excel file. Please make sure your input file is an Excel file (Excel 2007 is NOT supported.)') % input_excel.name)

        file_data = StringIO.StringIO()
        for chunk in input_excel.chunks():
            file_data.write(chunk)
        data['file_data'] = file_data.getvalue()

        try:
            xlrd.open_workbook(file_contents=data['file_data'])
        except xlrd.XLRDError, e:
            raise forms.ValidationError(_('Unable to open XLS file: %s' % e))

        return data


class ColumnAssignForm(forms.Form):
    def __init__(self, *args, **kwargs):
        if not 'modelform' in kwargs:
            raise KeyError('FieldSelectForm needs a modelform parameter')
        if not 'columns' in kwargs:
            raise KeyError('FieldSelectForm needs a columns parameter')
        self._modelform = kwargs['modelform']
        self._columns = kwargs['columns']
        del kwargs['modelform']
        del kwargs['columns']
        self.field_choices = []
        for name, field in self._modelform.fields.items():
            self.field_choices.append((name, field.label))
        self.field_choices.append((u'',u"don't use"))
        super(ColumnAssignForm, self).__init__(*args, **kwargs)
        for i, column in enumerate(self._columns):
            self.fields[str(i)] = forms.ChoiceField(choices=self.field_choices, required=False)

    def clean(self):
        for field, value in self.cleaned_data.items():
            if value == u'':
                # don't use this field
                del self.cleaned_data[field]
                continue
            for field2, value2 in self.cleaned_data.items():
                if field == field2:
                    continue
                if value == value2:
                    raise forms.ValidationError(_('Duplicated values'))
        return self.cleaned_data

    def get_excluded_fields(self):
        # todo: maybee whe should add some specific fields here
        fields = self.cleaned_data.values()
        return fields