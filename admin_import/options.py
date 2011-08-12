# xls import
from django.core.context_processors import csrf
from django.shortcuts import render
from admin_import.forms import XlsInputForm, ColumnAssignForm, create_partial_form

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
            PartialForm = create_partial_form(model_form, request.session['excel_import_excluded_fields'])
        if 'PartialForm' in locals() and request.method == 'POST' and '_send_common_data' in request.POST:
            partial_form = PartialForm(request.POST)
            if partial_form.is_valid():
                import_errors, import_count = do_import(sheet, model_form, request.session['excel_import_assignment'], partial_form.get_raw_data())
                if partial_form.cleaned_data['dry_run'] and not import_errors:
                    import_errors, import_count = do_import(sheet, model_form, request.session['excel_import_assignment'], partial_form.get_raw_data(),True)
        else:
            partial_form = PartialForm() if 'PartialForm' in locals() else None
        context.update({
            'sheet': sheet,
            'sheet_head':sheet_head,
            'column_assign_form': column_assign_form,
            'partial_form': partial_form,
            'import':{'errors':import_errors, 'count': import_count}if 'import_errors' in locals() else None,
        })
    context.update(csrf(request))
    return render(request, 'admin/excel_import/import_xls.html', context)


def add_import(admin):
    setattr(admin, 'import_xls_view', import_xls_view)
    setattr(admin, 'get_urls', decorate_get_urls(getattr(admin,'get_urls')))
    
def do_import(sheet, model_form, field_assignment, default_values, commit=False):
    errors = []
    count = 0
    for i in range(1,sheet.nrows):
        data = {v:sheet.cell(i,int(k)).value for k, v in field_assignment.items()}
        data.update(default_values)
        form = model_form(data)
        if form.is_valid():
            if commit:
                form.save()
                count +=1
        else:
            errors.append((sheet.row(i), form.errors))
    return errors, count