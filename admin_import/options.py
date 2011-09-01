# xls import
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
            request.session['excel_import_sheet'] = xlrd.open_workbook(file_contents=import_form.cleaned_data['input_excel'].read()).sheet_by_index(0)
    if 'import_form' not in locals():
        import_form = XlsInputForm()
    sheet = request.session.get('excel_import_sheet',None)
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
    for i in range(1,sheet.nrows):
        data = default_values.copy()
        for k, v in field_assignment.items():
            field = model_form().fields[v]
            value = sheet.cell(i,int(k)).value.strip()
            if hasattr(field, 'choices'):
                try:
                    value = dict(field.choices).values().index(value)
                except ValueError:
                    errors.append((sheet.row(i),ErrorDict(((v,ErrorList(["Could not assign value %s" % value])),))))
            data[v] = value
            # handle all choice fields
        form = model_form(data)
        if form.is_valid():
            if commit:
                form.save()
                count +=1
        else:
            errors.append((sheet.row(i), form.errors))
    return errors, count