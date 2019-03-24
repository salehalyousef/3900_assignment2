from django.contrib.auth.decorators import login_required
from django.shortcuts import render
from .models import *
from .forms import *
from django.shortcuts import render, get_object_or_404
from django.shortcuts import redirect
from django.db.models import Sum
import xlwt
from django.http import HttpResponse
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from .serializers import CustomerSerializer


now = timezone.now()
def home(request):
   return render(request, 'crm/home.html',
                 {'crm': home})

@login_required
def customer_list(request):
    customer = Customer.objects.filter(created_date__lte=timezone.now())
    return render(request, 'crm/customer_list.html',
                 {'customers': customer})


@login_required
def customer_edit(request, pk):
   customer = get_object_or_404(Customer, pk=pk)
   if request.method == "POST":
       # update
       form = CustomerForm(request.POST, instance=customer)
       if form.is_valid():
           customer = form.save(commit=False)
           customer.updated_date = timezone.now()
           customer.save()
           customer = Customer.objects.filter(created_date__lte=timezone.now())
           return render(request, 'crm/customer_list.html',
                         {'customers': customer})
   else:
        # edit
       form = CustomerForm(instance=customer)
   return render(request, 'crm/customer_edit.html', {'form': form})

@login_required
def customer_delete(request, pk):
   customer = get_object_or_404(Customer, pk=pk)
   customer.delete()
   return redirect('crm:customer_list')

@login_required
def service_list(request):
   services = Service.objects.filter(created_date__lte=timezone.now())
   return render(request, 'crm/service_list.html', {'services': services})


@login_required
def service_new(request):
   if request.method == "POST":
       form = ServiceForm(request.POST)
       if form.is_valid():
           service = form.save(commit=False)
           service.created_date = timezone.now()
           service.save()
           services = Service.objects.filter(created_date__lte=timezone.now())
           return render(request, 'crm/service_list.html',
                         {'services': services})
   else:
       form = ServiceForm()
       # print("Else")
   return render(request, 'crm/service_new.html', {'form': form})



@login_required
def service_edit(request, pk):
   service = get_object_or_404(Service, pk=pk)
   if request.method == "POST":
       form = ServiceForm(request.POST, instance=service)
       if form.is_valid():
           service = form.save()
           # service.customer = service.id
           service.updated_date = timezone.now()
           service.save()
           services = Service.objects.filter(created_date__lte=timezone.now())
           return render(request, 'crm/service_list.html', {'services': services})
   else:
       # print("else")
       form = ServiceForm(instance=service)
   return render(request, 'crm/service_edit.html', {'form': form})



@login_required
def summary(request, pk):
    customer = get_object_or_404(Customer, pk=pk)
    customers = Customer.objects.filter(created_date__lte=timezone.now())
    services = Service.objects.filter(cust_name=pk)
    products = Product.objects.filter(cust_name=pk)
    sum_service_charge = Service.objects.filter(cust_name=pk).aggregate(Sum('service_charge'))
    sum_product_charge = Product.objects.filter(cust_name=pk).aggregate(Sum('charge'))
    total_charge = sum_service_charge['service_charge__sum'] + sum_product_charge['charge__sum']
    print()
    return render(request, 'crm/summary.html', {'customers': customers,
                                                    'products': products,
                                                    'services': services,
                                                    'sum_service_charge': sum_service_charge,
                                                    'sum_product_charge': sum_product_charge,
                                                    'total_charge':total_charge,
                                                    'primary_key':pk})

@login_required
def service_delete(request, pk):
   service = get_object_or_404(Service, pk=pk)
   service.delete()
   return redirect('crm:service_list')


@login_required
def product_list(request):
   products = Product.objects.filter(created_date__lte=timezone.now())
   return render(request, 'crm/product_list.html', {'products': products})

@login_required
def product_new(request):
   if request.method == "POST":
       form = ProductForm(request.POST)
       if form.is_valid():
           product = form.save(commit=False)
           product.created_date = timezone.now()
           product.save()
           products = Product.objects.filter(created_date__lte=timezone.now())
           return render(request, 'crm/product_list.html',
                         {'products': products})
   else:
       form = ProductForm()
       # print("Else")
   return render(request, 'crm/product_new.html', {'form': form})

@login_required
def product_edit(request, pk):
   product = get_object_or_404(Product, pk=pk)
   if request.method == "POST":
       form = ProductForm(request.POST, instance=product)
       if form.is_valid():
           product = form.save()
           # service.customer = service.id
           product.updated_date = timezone.now()
           product.save()
           products = Product.objects.filter(created_date__lte=timezone.now())
           return render(request, 'crm/product_list.html', {'products': products})
   else:
       # print("else")
       form = ProductForm(instance=product)
   return render(request, 'crm/product_edit.html', {'form': form})


@login_required
def product_delete(request, pk):
   product = get_object_or_404(Product, pk=pk)
   product.delete()
   return redirect('crm:product_list')

def signup(request):
    if request.method == 'POST':
        form = UserSignUpForm(request.POST)
        if form.is_valid():
            user = form.save(commit=False)
            user.set_password(form.cleaned_data['password'])
            user.save()
            return render(request, 'registration/register_done.html', {'user': user})
    else:
        form = UserSignUpForm()
    return render(request,'registration/register.html',{'form': form})



def export_summary(request, pk):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="Summary.xls"'

    customer = get_object_or_404(Customer, pk=pk)
    customers = Customer.objects.filter(created_date__lte=timezone.now())
    services = Service.objects.filter(cust_name=pk)
    services_list = []
    products_list = []
    for service in services:
        mylist = []
        mylist.append(customer.cust_name)
        mylist.append(service.service_category)
        mylist.append(service.description)
        mylist.append(service.location)
        mylist.append(str(service.setup_time))
        mylist.append(str(service.cleanup_time))
        mylist.append(service.service_charge)
        services_list.append(mylist)
    products = Product.objects.filter(cust_name=pk)
    for product in products:
        mylist = []
        mylist.append(product.product)
        mylist.append(product.p_description)
        mylist.append(product.quantity)
        mylist.append(str(product.pickup_time))
        mylist.append(product.charge)
        products_list.append(mylist)
    sum_service_charge = Service.objects.filter(cust_name=pk).aggregate(Sum('service_charge'))
    sum_product_charge = Product.objects.filter(cust_name=pk).aggregate(Sum('charge'))
    total_charge = sum_service_charge['service_charge__sum'] + sum_product_charge['charge__sum']

    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Total')

    '''
        Writing Customer Summary in color
    '''
    row_num = 0
    font_style_header = xlwt.XFStyle()
    font_style_header.font.bold = True
    font_style_header.font.height = 480
    header_pattern = xlwt.Pattern()
    header_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    header_pattern.pattern_fore_colour = xlwt.Style.colour_map['ocean_blue']
    font_style_header.pattern = header_pattern

    ws.write(row_num, 0, "Customer Summary", font_style_header)
    ###########################################################################################
    ###########################################################################################
    # Sheet header
    row_num = 3
    font_style = xlwt.XFStyle()
    font_style.font.bold = True
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = xlwt.Style.colour_map['light_orange']
    font_style.pattern = pattern

    columns = ['Total of Service Charges and Product Charges' ]
    #Creating Header
    for col_num in range(len(columns)):
        ws.write(row_num, col_num, columns[col_num], font_style)

    # Sheet body, remaining rows
    font_style = xlwt.XFStyle()
    rows = []
    row_num = row_num + 1
    ws.write(row_num, 0, total_charge, xlwt.XFStyle())
    ###########################################################################################
    row_num = row_num + 3

    ws.write(row_num, 0, "Services Information", font_style_header)
    ###########################################################################################
    ###########################################################################################
    font_style.font.bold = True
    pattern.pattern_fore_colour = xlwt.Style.colour_map['light_orange']
    font_style.pattern = pattern
    row_num = 5 + row_num
    #Creating Header for Services
    columns = ['Customer Name', 'Service Category', 'Description', 'Location', 'Setup Time', 'Cleanup Time', 'Service Charge']
    for col_num in range(len(columns)):
        ws.write(row_num, col_num, columns[col_num], font_style)

    font_style = xlwt.XFStyle()

    for service in services_list:
        col_num = 0
        row_num = row_num + 1
        for data in service:
            ws.write(row_num, col_num, data, xlwt.XFStyle())
            col_num = col_num + 1

    row_num = row_num + 3
    font_style = xlwt.XFStyle()
    font_style.font.bold = True
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = xlwt.Style.colour_map['light_orange']
    font_style.pattern = pattern

    columns = ['Total of Service Charges' ]
    #Creating Header
    for col_num in range(len(columns)):
        ws.write(row_num, col_num, columns[col_num], font_style)

    # Sheet body, remaining rows
    font_style = xlwt.XFStyle()
    rows = []
    row_num = row_num + 1
    ws.write(row_num, 0, sum_service_charge['service_charge__sum'], xlwt.XFStyle())

    row_num = row_num + 3

    ws.write(row_num, 0, "Product Information", font_style_header)
    ###########################################################################################
    ###########################################################################################
    font_style.font.bold = True
    row_num = row_num + 5
    pattern.pattern_fore_colour = xlwt.Style.colour_map['light_orange']
    font_style.pattern = pattern
    #Creating Header for Services
    columns = ['Product', 'Description', 'Quantity', 'Pickup Time', 'Total Charge']
    for col_num in range(len(columns)):
        ws.write(row_num, col_num, columns[col_num], font_style)

    font_style = xlwt.XFStyle()
    for product in products_list:
        col_num = 0
        row_num = row_num + 1
        for data in product:
            ws.write(row_num, col_num, data, font_style)
            col_num = col_num + 1

    row_num = row_num + 3
    font_style = xlwt.XFStyle()
    font_style.font.bold = True
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = xlwt.Style.colour_map['light_orange']
    font_style.pattern = pattern

    columns = ['Total of Product Charges' ]
    #Creating Header
    for col_num in range(len(columns)):
        ws.write(row_num, col_num, columns[col_num], font_style)

    # Sheet body, remaining rows
    font_style = xlwt.XFStyle()
    rows = []
    row_num = row_num + 1
    ws.write(row_num, 0, sum_product_charge['charge__sum'], xlwt.XFStyle())
    wb.save(response)
    return response



# Lists all customers
class CustomerList(APIView):

    def get(self,request):
        customers_json = Customer.objects.all()
        serializer = CustomerSerializer(customers_json, many=True)
        return Response(serializer.data)
