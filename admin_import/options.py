from django import forms
from django.core.context_processors import csrf
from django.shortcuts import render
from django.forms.util import ErrorList, ErrorDict
from django.utils.translation import ugettext_lazy

from admin_import.forms import XlsInputForm, ColumnAssignForm

import xlrd

def decorate_get_urls(function):
    def wrapper(self):
        urls = function(self)
        from django.conf.urls.defaults import patterns
        export_urls = patterns('',
            (r'^import/$', self.admin_site.admin_view(self.import_xls_view))
        )
        return export_urls + urls
    return wrapper

def decorate_changelist_view(function):
    def wrapper(self, request, extra_context={}, **kwargs):
        extra_context.update({'has_import':True})
        return function(self, request, extra_context=extra_context, **kwargs)
    return wrapper

def import_xls_view(self, request):
    if request.method == 'POST' and '_send_file' in request.POST:
        # handle file and redirect
        import_form = XlsInputForm(request.POST, request.FILES)
        if import_form.is_valid():
            request.session['excel_import_sheet'] = import_form.cleaned_data['file_data']

    if 'import_form' not in locals():
        import_form = XlsInputForm()

    sheet = request.session.get('excel_import_sheet', None)
    if sheet:
        sheet = xlrd.open_workbook(file_contents=sheet).sheet_by_index(0)

    context = {
        'import_form':import_form,
        'app_label': self.model._meta.app_label,
        'opts': self.model._meta,
    }
    if sheet:
        sheet_head = (sheet.row(i) for i in range(min(3, sheet.nrows)))
        model_form = self.get_form(request)
        form_instance = model_form()
        columns = (field.value for field in sheet.row(0))
        if request.method == 'POST' and '_send_assignment' in request.POST:
            column_assign_form = ColumnAssignForm(request.POST, modelform=form_instance, columns=columns)
            if column_assign_form.is_valid():
                request.session['excel_import_excluded_fields'] = column_assign_form.get_excluded_fields()
                request.session['excel_import_assignment'] = column_assign_form.cleaned_data
        elif 'excel_import_assignment' in request.session:
            column_assign_form = ColumnAssignForm(request.session['excel_import_assignment'], modelform=form_instance, columns=columns)
        else:
            column_assign_form = ColumnAssignForm(modelform=form_instance, columns=columns)
        if 'excel_import_excluded_fields' in request.session:
            PartialForm = self.get_form(request, exclude=request.session['excel_import_excluded_fields'])
            PartialForm.base_fields['dry_run'] = forms.BooleanField(label=ugettext_lazy('Dry run'),
                help_text=ugettext_lazy('Uncheck this checkbox if you actually want to save the data in the database!'),
                required=False, initial=True)

        if 'PartialForm' in locals() and request.method == 'POST' and '_send_common_data' in request.POST:
            partial_form = PartialForm(request.POST)
            if partial_form.is_valid():
                import_errors, import_count = do_import(sheet, model_form, request.session['excel_import_assignment'], request.POST)
                if not partial_form.cleaned_data['dry_run'] and not import_errors:
                    import_errors, import_count = do_import(sheet, model_form, request.session['excel_import_assignment'], request.POST, True)
                context.update({'import':{'errors':import_errors,
                                          'count': import_count,
                                          'dry_run': partial_form.cleaned_data['dry_run'],}})
        else:
            partial_form = PartialForm() if 'PartialForm' in locals() else None
        context.update({
            'sheet': sheet,
            'sheet_head':sheet_head,
            'column_assign_form': column_assign_form,
            'partial_form': partial_form,
        })
    context.update(csrf(request))
    return render(request, 'admin/excel_import/import_xls.html', context)


def add_import(admin, add_button=False):
    setattr(admin, 'import_xls_view', import_xls_view)
    setattr(admin, 'get_urls', decorate_get_urls(getattr(admin,'get_urls')))
    if add_button:
        setattr(admin, 'changelist_view', decorate_changelist_view(getattr(admin, 'changelist_view')))
        setattr(admin, 'change_list_template', 'admin/excel_import/changelist_view.html')

def do_import(sheet, model_form, field_assignment, default_values, commit=False):
    errors = []
    count = 0
    form_instance = model_form()

    # Prepare choices mappings to handle choice fields
    forward_choices = {}
    reverse_choices = {}
    for k, v in field_assignment.items():
        field = form_instance.fields[v]
        if hasattr(field, 'choices'):
            forward_choices[k] = choice_dict = dict(field.choices)
            reverse_choices[k] = dict((r[1], r[0]) for r in choice_dict.items())

    for i in range(1, sheet.nrows):
        data = default_values.copy()
        print 'Processing %s/%s: %s' % (i, sheet.nrows, data)
        values = sheet.row_values(i)

        for k, v in field_assignment.items():
            field = form_instance.fields[v]
            value = values[int(k)]

            # Normalize values a little bit -- this is necessary because when
            # reading from the excel file, we only get a subset of the types
            # Django itself supports through its forms and models.
            if isinstance(value, float):
                if value - int(value) == 0.0:
                    value = int(value)
            elif isinstance(value, basestring):
                value = value.strip()

            if k in forward_choices:
                if value in forward_choices[k]:
                    pass # Ok.
                else:
                    if value in reverse_choices[k]:
                        value = reverse_choices[k][value]
                    else:
                        errors.append((
                            values,
                            ErrorDict((
                                (v, ErrorList(["Could not assign value %s" % value])),
                                ))
                            ))

            data[v] = value

        form = model_form(data)
        if form.is_valid():
            if commit:
                form.save()
                count +=1
        else:
            errors.append((sheet.row(i), form.errors))
    return errors, count