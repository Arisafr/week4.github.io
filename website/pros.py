import os
from flask import Blueprint, flash, make_response, redirect, render_template, request, url_for, send_file
from numpy import append
from werkzeug.utils import secure_filename
import pandas as pd
from flask_wtf import FlaskForm
from wtforms import SubmitField, FileField
from .views import FileUploadForm

# pros = proses
pros = Blueprint('pros',__name__)
filename = None 
# Global variable to store the dataframe
df = pd.DataFrame()

@pros.route('/home', methods=['POST'])
def upload():
    form = FileUploadForm()
    if form.validate_on_submit():
        file = request.files['file']
        if file.filename == '':
            message = "File Belum Dipilih"
        else:
            filename = secure_filename(file.filename)
            file.save(filename)
            global df
            df = pd.read_excel(filename)
            message = "File Uploaded " + filename + "\n\nPastikan proses yang dikirim sesuai dengan file yang diupload"
    else:
        message = None # add this line
    print (df)
    return render_template('home.html', form=form, message=message) # pass the form variable to the template


@pros.route("/reportDetail")
def reportDetail():
    # kode untuk memproses laporan detail
    import pandas as pd
    global df

    data = df

    data.drop('Corporate ID', axis=1, inplace=True)
    data.drop('First Name', axis=1, inplace=True)
    data.drop('Last Name', axis=1, inplace=True)
    data.drop('Organization', axis=1, inplace=True)
    data.drop('Assignee', axis=1, inplace=True)
    data.drop('Phone Number', axis=1, inplace=True)
    data.drop('Site', axis=1, inplace=True)
    data.drop('Summary', axis=1, inplace=True)
    data.drop('Operational Categorization 1', axis=1, inplace=True)
    data.drop('Operational Categorization 2', axis=1, inplace=True)
    data.drop('Operational Categorization 3', axis=1, inplace=True)
    data.drop('Product Category 2', axis=1, inplace=True)
    data.drop('Product Category 3', axis=1, inplace=True)
    data.drop('Resolution Product Category 1', axis=1, inplace=True)
    data.drop('Resolution Product Category 2', axis=1, inplace=True)
    data.drop('Resolution Product Category 3', axis=1, inplace=True)
    data = data[data['Product Category 1'] != 'TRUESIGHT EVENTS']
    data = data.drop(data[data['Reported Source'] == 'BMC Impact Manager Event'].index)

    """## Data Labeling
    ### Ticket TI
    Proses
    """
    count = 0
    CountSelesai = 0
    CountBelum = 0
    for index, row in data.iterrows():
        # mengambil nilai pada kolom tertentu
        nilai_kolom = row['Status']
        # melakukan operasi atau pemrosesan data pada nilai_kolom
        # ...
        if nilai_kolom in ['Closed', 'Cancelled', 'Resolved']:
            CountSelesai += 1
        else:
            CountBelum += 1

    TotalInsiden = data['Incident Number'].count()
    SelesaiDitangani = TotalInsiden - data['Closed Date'].isnull().sum()
    BelumSelesaiDitangani = data['Closed Date'].isnull().sum()
    PresentaseSudah = SelesaiDitangani/TotalInsiden*100
    PresentaseBelum = BelumSelesaiDitangani/TotalInsiden*100

    """Hasil"""

    #@title Fungsi Cetak
    def TicketPrint():
        Ticket = pd.DataFrame({'Status': ['Selesai', 'Belum Selesai'], 'Count': [CountSelesai, CountBelum],'Presentase' :[str("{:.2f}".format(PresentaseSudah)) + '%', str("{:.2f}".format(PresentaseBelum)) + '%']})
        print('\n\nTicketing')
        print(Ticket)

    ### Priority

    # Checked

    #@title Cetak data apa yang ada di dalam kolom
    unique_categories = data['Priority'].unique()
    Prior = []

    for category in unique_categories:
        Prior.append(category)

    """Hasil"""

    CountCriticalSelesai = 0
    CountCriticalBelum = 0
    CountHighSelesai = 0
    CountHighBelum = 0
    CountMediumSelesai = 0
    CountMediumBelum = 0
    CountLowSelesai = 0
    CountLowBelum = 0

    for index, row in data.iterrows():
        nilai_status = row['Status']
        nilai_priority = row['Priority']
        if nilai_status in ['Closed', 'Cancelled', 'Resolved']:
            if nilai_priority == 'Critical':
                CountCriticalSelesai += 1
            elif nilai_priority == 'High':
                CountHighSelesai += 1
            elif nilai_priority == 'Medium':
                CountMediumSelesai += 1
            elif nilai_priority == 'Low':
                CountLowSelesai += 1
        else:
            if nilai_priority == 'Critical':
                CountCriticalBelum += 1
            elif nilai_priority == 'High':
                CountHighBelum += 1
            elif nilai_priority == 'Medium':
                CountMediumBelum += 1
            elif nilai_priority == 'Low':
                CountLowBelum += 1

    Reported_Priority = pd.DataFrame({
        'Priority': ['Critical', 'High', 'Medium', 'Low'], 
        'Selesai': [CountCriticalSelesai, CountHighSelesai, CountMediumSelesai, CountLowSelesai],
        'Belum Selesai': [CountCriticalBelum, CountHighBelum, CountMediumBelum, CountLowBelum]
    })

    Reported_Priority = Reported_Priority.set_index('Priority')

    def ReportPriority():
        print('\n\nReported Priority')
        print(Reported_Priority)

    """### Reported Sources

    Parameter
    """

    #@title Cetak data apa yang ada di dalam kolom
    unique_categories = data['Reported Source'].unique()

    #@title Cetak data menjadi array
    unique_categories = data['Reported Source'].unique()
    Reported_Sources = []

    for category in unique_categories:
        Reported_Sources.append(category)

    """Hasil"""

    # Ambil kolom Reported Source dan hitung jumlah kategori
    kategori = data['Reported Source'].value_counts()

    # Buat DataFrame baru dengan nama Reported Sources dan urutkan secara menurun
    Reported_Sources = pd.DataFrame(kategori).reset_index().rename(columns={'index': 'Reported Source', 'Reported Source': 'Count'})

    # Tambahkan kolom presentase
    percentage = Reported_Sources['Count']/Reported_Sources['Count'].sum()*100
    Reported_Sources['Percentage'] = percentage.map('{:.2f}%'.format)

    # Urutkan secara menurun berdasarkan Count
    Reported_Sources = Reported_Sources.sort_values(by='Count', ascending=False)

    def ReportedSources():
        print ('\n\n')
        print ('Reported Sources')
        print(Reported_Sources)

    """### Product

    Parameter
    """

    #@title ubah data kolom x menjadi array
    unique_categories = data['Product Category 1'].unique()
    Product_Category = []

    for products in unique_categories:
        Product_Category.append(products)
    
    """Hasil"""

    # Ambil kolom Reported Source dan hitung jumlah kategori
    kategori = data['Product Category 1'].value_counts()

    # Buat DataFrame baru dengan nama Reported Sources
    Product_Category = pd.DataFrame(kategori)

    # Tambahkan kolom "Count" yang merupakan jumlah dari setiap kategori
    Product_Category['Count'] = Product_Category['Product Category 1']

    # Tambahkan kolom "Percentage" yang merupakan presentase dari setiap kategori
    Product_Category['Percentage'] = Product_Category['Count'] / Product_Category['Count'].sum() * 100
    Product_Category['Percentage'] = Product_Category['Percentage'].map('{:.2f}%'.format)

    # Urutkan dataframe secara menurun berdasarkan kolom "Count"
    Product_Category = Product_Category.sort_values('Count', ascending=False)

    # Pilih 5 data teratas
    top_4 = Product_Category.head(4)

    # Hitung jumlah dari data di luar 5 data teratas
    others_count = Product_Category['Count'].iloc[4:].sum()

    # Buat dataframe "Lainnya" dengan jumlah data di luar 5 data teratas
    others = pd.DataFrame({
        'Product Category 1': 'Lainnya',
        'Count': others_count,
        'Percentage': '{:.2f}%'.format(others_count / Product_Category['Count'].sum() * 100)
    }, index=['Lainnya'])

    # Gabungkan 5 data teratas dan dataframe "Lainnya"
    Product_Category = pd.concat([top_4, others])

    # Tampilkan kolom "Product Category 1", "Count", dan "Percentage"
    def ProductCategory():
        print('\n\nProduct Category')
        print(Product_Category[['Count', 'Percentage']].rename_axis('Product Category 1'))

    """### Assigned Group

    ### Parameter
    """

    #@title buat isi kolom x menjadi array
    unique_categories = data['Assigned Group'].unique()
    assignedGroup = []

    for category in unique_categories:
        assignedGroup.append(category)

    """Hasil"""

    from collections import defaultdict
    import pandas as pd

    # Membuat default dictionary untuk assignedGroup
    assignedGroup = defaultdict(lambda: {'Selesai': 0, 'Belum': 0})

    # Menghitung jumlah tiket yang Selesai dan Belum pada setiap Assigned Group
    for group, data_group in data.groupby('Assigned Group'):
        for status in ['Closed', 'Cancelled', 'Resolved']:
            assignedGroup[group]['Selesai'] += (data_group['Status'] == status).sum()
        assignedGroup[group]['Belum'] = len(data_group) - assignedGroup[group]['Selesai']

    # Membuat dataframe dari dictionary assignedGroup dan mengurutkannya berdasarkan kolom "Selesai"
    df = pd.DataFrame.from_dict(assignedGroup, orient='index').sort_values('Selesai', ascending=False)

    # Menampilkan hanya 5 data teratas dan sisanya dijumlahkan dan dimasukkan ke dalam kategori 'Lainnya'
    top_5 = df.head(4)
    others = pd.DataFrame(df.iloc[4:].sum()).T.rename(index={0: 'Lainnya'})
    df_new = pd.concat([top_5, others])

    # Menampilkan kolom Assigned Group, Selesai, dan Belum yang telah diurutkan
    def Assigned():
        print ('\n\nAssigned Group')
        print(df_new[['Selesai', 'Belum']].rename_axis('AssignedGroup'))
        # Menyimpan hasil output ke dalam file excel
        df_new[['Selesai', 'Belum']].to_excel('AssignedGroup.xlsx')

    # Total Of Data
    print ('Total Of Data')
    print(data.shape)

    # Memasukkan hasil print ke dalam DataFrame
    ticket = pd.DataFrame({'Status': ['Selesai', 'Belum Selesai'], 'Count': [CountSelesai, CountBelum],'Presentase' :[str("{:.2f}".format(PresentaseSudah)) + '%', str("{:.2f}".format(PresentaseBelum)) + '%']})
    df_priority = pd.DataFrame({
        'Priority': ['Critical', 'High', 'Medium', 'Low'], 
        'Selesai': [CountCriticalSelesai, CountHighSelesai, CountMediumSelesai, CountLowSelesai],
        'Belum Selesai': [CountCriticalBelum, CountHighBelum, CountMediumBelum, CountLowBelum]
    })
    df_assigned = df_new[['Selesai', 'Belum']]
    df_product = Product_Category[['Count', 'Percentage']]
    df_source = Reported_Sources

    # writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    with pd.ExcelWriter('Hasil_DLH.xlsx') as writer:
        ticket.to_excel(writer, sheet_name='Ticketing')
        df_assigned[['Selesai', 'Belum']].rename_axis('AssignedGroup').to_excel(writer, sheet_name='Assigned Group')
        df_product.to_excel(writer, sheet_name='Product Category')
        df_priority.to_excel(writer, sheet_name='Reported Priority')
        df_source.to_excel(writer, sheet_name='Reported Sources')

    global filename
    filename = 'Hasil_DLH.xlsx'

    return "Laporan Detail sudah diproses!"

@pros.route("/workOrder")
def WO():
    from collections import defaultdict
    import pandas as pd

    writer = None
    global df
    data = df
    print(data.shape)


    """## PreProcessing
    """

    #@title Hapus kolom dibawah
    data.drop('Notes', axis=1, inplace=True)
    data.drop('Requested By First Name', axis=1, inplace=True)
    data.drop('Requested by Last Name', axis=1, inplace=True)
    data.drop('Request Manager Support Group', axis=1, inplace=True)
    data.drop('Request Assignee', axis=1, inplace=True)
    data.drop('On Site', axis=1, inplace=True)
    data.drop('Service Name', axis=1, inplace=True)
    data.drop('Scheduled Start Date', axis=1, inplace=True)
    data.drop('Scheduled End Date', axis=1, inplace=True)
    data.drop('Actual Start Date', axis=1, inplace=True)
    data.drop('Actual End Date', axis=1, inplace=True)
    data.drop('Completed Date__c', axis=1, inplace=True)
    data.drop('Status Reason', axis=1, inplace=True)
    data.drop('Submit Date', axis=1, inplace=True)
    data.drop('Request Manager', axis=1, inplace=True)
    data.drop('Resolution', axis=1, inplace=True)
    data.drop('Level', axis=1, inplace=True)

    """## Data Labeling

    ## Ticket TI
    """

    count = 0
    totaldata = data['Status'].count()
    CountSelesai = 0
    CountBelum = 0

    for index, row in data.iterrows():
        # mengambil nilai pada kolom tertentu
        nilai_kolom = row['Status']
        # melakukan operasi atau pemrosesan data pada nilai_kolom
        # ...
        if nilai_kolom in ['Closed', 'Cancelled', 'Resolved']:
            CountSelesai += 1
        else:
            CountBelum += 1

    Total = (CountBelum+CountSelesai)
    PresentaseSudah = CountSelesai/Total*100
    PresentaseBelum = CountBelum/Total*100

    #@title Fungsi Print Tiket TI
    def TicketPrint():
        Ticket = pd.DataFrame({'Status': ['Selesai', 'Belum Selesai'], 'Count': [CountSelesai, CountBelum],'Presentase' :[str("{:.2f}".format(PresentaseSudah)) + '%', str("{:.2f}".format(PresentaseBelum)) + '%']})
        print('\n\nTicketing')
        print(Ticket)


    """## Priority
    ### Checked
    """

    #@title Pengecekan isi kolom X
    unique_categories = data['Priority'].unique()
    Prior = []

    for category in unique_categories:
        Prior.append(category)

    #@title Pengecekan isi kolom X
    unique_categories = data['Status'].unique()

    """### Hasil"""

    CountCriticalSelesai = 0
    CountCriticalBelum = 0
    CountHighSelesai = 0
    CountHighBelum = 0
    CountMediumSelesai = 0
    CountMediumBelum = 0
    CountLowSelesai = 0
    CountLowBelum = 0

    for index, row in data.iterrows():
        nilai_status = row['Status']
        nilai_priority = row['Priority']
        if nilai_status in ['Closed', 'Cancelled', 'Resolved']:
            if nilai_priority == 'Critical':
                CountCriticalSelesai += 1
            elif nilai_priority == 'High':
                CountHighSelesai += 1
            elif nilai_priority == 'Medium':
                CountMediumSelesai += 1
            elif nilai_priority == 'Low':
                CountLowSelesai += 1
        else:
            if nilai_priority == 'Critical':
                CountCriticalBelum += 1
            elif nilai_priority == 'High':
                CountHighBelum += 1
            elif nilai_priority == 'Medium':
                CountMediumBelum += 1
            elif nilai_priority == 'Low':
                CountLowBelum += 1

    Reported_Priority = pd.DataFrame({
        'Priority': ['Critical', 'High', 'Medium', 'Low'], 
        'Selesai': [CountCriticalSelesai, CountHighSelesai, CountMediumSelesai, CountLowSelesai],
        'Belum Selesai': [CountCriticalBelum, CountHighBelum, CountMediumBelum, CountLowBelum]
    })

    def ReportPriority():
        print('\n\nReported Priority')
        print(Reported_Priority)


    """### Assigned Group

    ### Parameter
    """

    #@title buat isi kolom x menjadi array
    unique_categories = data['Request Assignee Support Group'].unique()
    assignedGroup = []

    for category in unique_categories:
        assignedGroup.append(category)

    """Hasil"""

    from collections import defaultdict
    import pandas as pd

    # Membuat default dictionary untuk assignedGroup
    assignedGroup = defaultdict(lambda: {'Selesai': 0, 'Belum': 0})

    # Menghitung jumlah tiket yang Selesai dan Belum pada setiap Assigned Group
    for group, data_group in data.groupby('Request Assignee Support Group'):
        for status in ['Closed', 'Cancelled', 'Resolved']:
            assignedGroup[group]['Selesai'] += (data_group['Status'] == status).sum()
        assignedGroup[group]['Belum'] = len(data_group) - assignedGroup[group]['Selesai']

    # Membuat dataframe dari dictionary assignedGroup dan mengurutkannya berdasarkan kolom "Selesai"
    df = pd.DataFrame.from_dict(assignedGroup, orient='index').sort_values('Selesai', ascending=False)

    # Menampilkan hanya 5 data teratas dan sisanya dijumlahkan dan dimasukkan ke dalam kategori 'Lainnya'
    top_5 = df.head(4)
    others = pd.DataFrame(df.iloc[4:].sum()).T.rename(index={0: 'Lainnya'})
    df_new = pd.concat([top_5, others])

    # Menampilkan kolom Assigned Group, Selesai, dan Belum yang telah diurutkan
    def Assigned():
        print ('\n\nAssigned Group')
        print(df_new[['Selesai', 'Belum']].rename_axis('AssignedGroup'))
        # Menyimpan hasil output ke dalam file excel
        df_new[['Selesai', 'Belum']].to_excel('AssignedGroup.xlsx')


    """## Kategori Produk

    ### Parameter
    """
    #@title Pengubahan data kolom menjadi Array untuk di proses
    unique_categories = data['Prod Cat 1'].unique()
    Product_Category = []

    for products in unique_categories:
        Product_Category.append(products)


    """### Hasil"""

    # Ambil kolom Product Cat 1 dan hitung jumlah kategori
    kategori = data['Prod Cat 1'].value_counts()

    # Buat DataFrame baru dengan nama Product Category dan urutkan secara menurun
    Product_Category = pd.DataFrame(kategori).reset_index().rename(columns={'index': 'Product Category 1', 'Prod Cat 1': 'Count'})

    # Tambahkan kolom presentase
    percentage = Product_Category['Count']/Product_Category['Count'].sum()*100
    Product_Category['Percentage'] = percentage.map('{:.2f}%'.format)

    # Urutkan secara menurun berdasarkan Count
    Product_Category = Product_Category.sort_values(by='Count', ascending=False)

    # Pilih 4 data teratas
    top_4 = Product_Category.head(4)

    # Hitung jumlah dari data di luar 4 data teratas
    others_count = Product_Category['Count'].iloc[4:].sum()

    # Buat dataframe "Lainnya" dengan jumlah data di luar 4 data teratas
    others = pd.DataFrame({
        'Product Category 1': 'Lainnya',
        'Count': others_count,
        'Percentage': '{:.2f}%'.format(others_count / Product_Category['Count'].sum() * 100)
    }, index=['Lainnya'])

    # Gabungkan 4 data teratas dan dataframe "Lainnya"
    Product_Category = pd.concat([top_4, others])

    def ProductCategory():
        print('\n\nProduct Category')
        print(Product_Category[['Product Category 1', 'Count', 'Percentage']].rename_axis('No'))

    ProductCategory()

    """## Lokasi"""
    #@title Pengubahan data pada kolom menjadi array
    unique_categories = data['Request Assignee Support Organization'].unique()
    AssigneeSupportOrganization = []

    for organization in unique_categories:
        AssigneeSupportOrganization.append(organization)

    # Ambil kolom Request Assignee Support Organization dan hitung jumlah kategori
    kategori = data['Request Assignee Support Organization'].value_counts()

    # Buat DataFrame baru dengan nama Product Category dan urutkan secara menurun
    AssigneeSupportOrganization = pd.DataFrame(kategori).reset_index().rename(columns={'index': 'Support Group', 'Request Assignee Support Organization': 'Count'})

    # Tambahkan kolom presentase
    percentage = AssigneeSupportOrganization['Count']/Product_Category['Count'].sum()*100
    AssigneeSupportOrganization['Percentage'] = percentage.map('{:.2f}%'.format)

    # Urutkan secara menurun berdasarkan Count
    AssigneeSupportOrganization = AssigneeSupportOrganization.sort_values(by='Count', ascending=False)

    def Assignee_Support_Organization():
        print('\n\n')
        print('Lokasi')
        print(AssigneeSupportOrganization)
        Assignee_Support_Organization()


    """## Cetak Ke dalam Excel"""
    Ticket = pd.DataFrame({'Status': ['Selesai', 'Belum Selesai'], 'Count': [CountSelesai, CountBelum],'Presentase' :[str("{:.2f}".format(PresentaseSudah)) + '%', str("{:.2f}".format(PresentaseBelum)) + '%']})

    def ReportPriority():
        print('\n\nReported Priority')
        print(Reported_Priority)

    def Assigned():
        print ('\n\nAssigned Group')
        print(df_new[['Selesai', 'Belum']].rename_axis('Assigned Group'))
    df_assigned = df_new[['Selesai', 'Belum']]

    def ProductCategory():
        print('\n\n')
        print('Product Category')
        print(Product_Category)


    def Assignee_Support_Organization():
        print('\n\n')
        print('Lokasi')
        print(AssigneeSupportOrganization)

    global filename
    filename = 'Hasil_WO.xlsx'

    with pd.ExcelWriter('Hasil_WO.xlsx') as writer:
        Ticket.to_excel(writer, sheet_name='Ticketing', index=False)
        Reported_Priority.to_excel(writer, sheet_name='Reported Priority', index=False)
        df_assigned[['Selesai', 'Belum']].rename_axis('AssignedGroup').to_excel(writer, sheet_name='Assigned Group')
        Product_Category.to_excel(writer, sheet_name='Product Category', index=False)
        AssigneeSupportOrganization.to_excel(writer, sheet_name='Assignee Support Organization', index=False)

    return "Laporan Work Order sudah diproses!"

@pros.route('/sla')    
def SLA():
    global df
    # Menghapus baris dengan nilai null
    df.dropna(inplace=True)

    """## Measurement Status"""

    unique_categories = df['MeasurementStatus'].unique()

    Met = 0
    Missed = 0

    for index, row in df.iterrows():
        nilai_status = row['MeasurementStatus']
        if nilai_status in ['Met']:
            Met = Met + 1
        else:
            Missed = Missed + 1

    Measurement_Status = pd.DataFrame({
        'Measurement Status': ['Met', 'Missed'], 
        'Count': [Met, Missed]
    })

    Measurement_Status = Measurement_Status.set_index('Measurement Status')
    # print('\nMeasurement Status')
    print(Measurement_Status)


    """## Assigned Group

    ### Parameter
    """
    unique_categories = df['Assigned Group'].unique()
    assignedGroup = []

    for category in unique_categories:
        assignedGroup.append(category)
    
    from collections import defaultdict

    assignedGroup = defaultdict(lambda: {'Met': 0, 'Missed': 0})

    for group, data_group in df.groupby('Assigned Group'):
        for status in ['Met']:
            assignedGroup[group]['Met'] += (data_group['MeasurementStatus'] == status).sum()
        assignedGroup[group]['Missed'] = len(data_group) - assignedGroup[group]['Met']


    # Membuat dataframe dari dictionary assignedGroup
    df = pd.DataFrame.from_dict(assignedGroup, orient='index')
    df = df.sort_values('Met', ascending=False)

    top_5 = df.head(4)
    others = pd.DataFrame(df.iloc[4:].sum()).T.rename(index={0: 'Lainnya'})
    df_new = pd.concat([top_5, others])

    # Menampilkan kolom "Assigned Group", "Met", "Missed", dan "Total"
    print('\nAssigned Group')
    print(df_new[['Met', 'Missed']].rename_axis('AssignedGroup'))
    df_assigned = df_new[['Met', 'Missed']]

    # membuat file excel
    with pd.ExcelWriter('HasilReport_SLA_Incident.xlsx') as writer:
    # menyimpan data ke dalam sheet Measurement Status
        Measurement_Status.to_excel(writer, sheet_name='Measurement Status')
    # menyimpan data ke dalam sheet Assigned Group
        df_assigned.rename_axis('AssignedGroup').to_excel(writer, sheet_name='Assigned Group')
    global filename
    filename = 'HasilReport_SLA_Incident.xlsx'
    return 'Laporan Detail sudah diproses!'


@pros.route('/download', methods=['POST'])
def download():
    global filename
    print (df)
    if filename is not None:
        file = filename
        # menghasilkan respons HTTP
        response = make_response()
        response.data = open(file, 'rb').read()
        response.headers.set('Content-Disposition', 'attachment', filename=file)
        response.headers.set('Content-Type', 'application/vnd.ms-excel')
        return response
    else:
        return 'File tidak ditemukan!'
