import os
import json
import requests
import mimetypes
import numbers
import xlrd
import xlsxwriter
from datetime import datetime,timedelta
from dateutil.parser import parse
from dateutil import parser
from string import whitespace
from collections import defaultdict
#import teradata as td
from werkzeug.utils import secure_filename
from werkzeug.debug import get_current_traceback
#from flask_track_usage import TrackUsage
#from flask_track_usage.storage import Storage

from flask import Flask, render_template, request, jsonify, redirect, \
                  url_for, session, flash, make_response, abort, send_from_directory

#import oracle_connect
import config
#import td_con
from webbrowser import Opera


UPLOAD_FOLDER = 'D:/POC'
ALLOWED_EXTENSIONS = set(['txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif','xls','xlsx','doc','docs'])


app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['TRACK_USAGE_INCLUDE_OR_EXCLUDE_VIEWS'] = 'include'
app.secret_key = "development-key"

class Usage(Storage):
    def store(self, data):
        oracle_connect.track_user_details(data, session)
        
track = TrackUsage(app, Usage())

@app.context_processor
def custom_context_processor():
    return dict(user_name = session['userName'])

@app.before_request
def userSession():
    try:
        if 'userName' not in session:
            session.permanent = False
            userName = config.TEST_USER_NAME
            bemsId = config.TEST_BEMSID
            emailId = "test@boeing.com"
            if config.WITH_WSSO:
                # get user info from WSSO
                userName = request.headers.get("boeingdisplayname")
                bemsId = request.headers.get("boeingbemsid")
                emailId = request.headers.get("mail")
            session['userName'] = str(userName)
            session['bemsId'] = int(bemsId)
            session['emailId'] = emailId
			
    except:
        track= get_current_traceback(skip=1, show_hidden_frames=True,
            ignore_system_exceptions=False)
        track.log()
        abort(500)


@app.route('/')
def app_home():
    return redirect(url_for('index'))

@track.include
@app.route('/home')
def index():
    return render_template('home.html')

@track.include
@app.route('/sales')
def sales():
    '''result = dict()
    result = td_con.get_column_datatype()
    print (result)'''
    '''field_to_display_sales = oracle_connect.get_fields_to_display_sales()
    select_criteria_sales = oracle_connect.get_selection_criteria_sales()
    group_by_sales = oracle_connect.get_group_by_sales()
    file_name_sales = oracle_connect.get_file_name_cust(user_name = session['userName'])
    return render_template('sales.html',
                           field_to_display_sales=field_to_display_sales,
                           select_criteria_sales=select_criteria_sales,
                           file_name_sales=file_name_sales,
                           group_by_sales=group_by_sales)'''
    return render_template('sales.html')

@track.include
@app.route('/parts')
def parts():
    #field_to_display_parts = td_con.get_fields_to_display_parts()
    '''field_to_display_parts = oracle_connect.get_fields_to_display_parts()
    select_criteria_parts = oracle_connect.get_selection_criteria_parts()
    group_by_parts = oracle_connect.get_group_by_parts()
    file_name_parts = oracle_connect.get_file_name_parts(user_name = session['userName'])
    return render_template('parts.html',
                           field_to_display_parts=field_to_display_parts,
                           select_criteria_parts=select_criteria_parts,
                           file_name_parts=file_name_parts,
                           group_by_parts=group_by_parts)'''
    return render_template('parts.html')

@app.route('/upload')
def upload():
    return render_template('upload.html')
	
@app.route('/downloadTemplate/<path:xls_name>')
def download_template(xls_name):
    """ return the template"""
    template_map = {"customer": "Customer_list.xlsx",
                    "parts": "Part_list.xlsx"}
    return send_from_directory('static', template_map.get(xls_name,"service_template.xlsx"), as_attachment=True)



@app.route('/adhocQuery/<string:data>', methods=['POST'])
def adhoc_query(data):
    result = dict() 
    result_query = ''
    searched_column_list = ''
    result_query, searched_column_list = form_query(data, request)

    result['data'] = td_con.get_data_from_db(result_query, searched_column_list)
    result['message'] = 'success'

    return jsonify(result)

@app.route('/saveQuery/<string:data>', methods=['POST'])
def save_query(data):
    result = dict()
    query_details = {}
    query, searched_column_list = form_query(data, request)
    print('*************************************************')
    query_details['BEMSID'] = session.get('bemsId', '1111')
    if data == 'sales':
        query_details['QUERY_NAME'] =  request.form['SalesQueryName']
        query_details['QUERY_DESC'] =  request.form['SalesQueryDesc']
    else:
        query_details['QUERY_NAME'] =  request.form['PartsQueryName']
        query_details['QUERY_DESC'] =  request.form['PartsQueryDesc']
    query_details['QUERY_TYPE'] = data
    query_details['QUERY'] = query
    print(query_details)
    oracle_connect.save_query_text(query_details)
    return 'success'

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'],filename)
	
@app.route('/UploadFile', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        finalData=[]
        result = {}
        files = request.files.get('file')
        file_type = request.form['file_name']
        file_details = request.form['fileDetails']
        file_name = files.filename
        filename, file_extension = os.path.splitext(file_name)
 
        #oracle query with distinc file name and then compare with user filename
        user_existing_file_name = oracle_connect.check_user_file_name(user_name=session['userName'], file_name=file_details, file_type = file_type)
        if user_existing_file_name:
            return "exists"
        else:
            if file_extension.lower() not in ['.xlsx','.xls']:
                message="Invalid file, file should be of type .xls, .xlsx"
                return emptyData,message
            workbook = xlrd.open_workbook(file_contents=files.read())

            sheet = workbook.sheet_by_index(0)
            
            for j in range(2, sheet.nrows):
                data = {'USER_NAME': session['userName'], 'FILE_NAME':file_details}
                col_header = sheet.cell(1,0).value
                cell_value = sheet.cell(j,0).value
                data[col_header] = cell_value
                finalData.append(data)
            oracle_connect.insert_excel_data_to_db(list_of_dicts=finalData, file_type=file_type)
            return "success"

@app.route('/getPartList', methods=['POST'])
def get_part_list():
    file_name = oracle_connect.get_file_name_parts(session['userName'])
    return jsonify(file_name)

@app.route('/getCustList', methods=['POST'])
def get_cust_list():
    file_name = oracle_connect.get_file_name_cust(session['userName'])
    return jsonify(file_name)

def form_query(data, request):
    if request.method == 'POST':
        #import pdb;pdb.set_trace();
        query = ''
        cust_list = ''
        file_field = ''
        fileList = ''
        part_list= ' '
        data_list = ''
        grp_flag = False
        groupByField = ''
        aggr_func = ''
        grp_by_query = ''
        user_name = config.USER_NAME
        search_to = request.form.get('searchField',None)
        search_to_list = search_to.split(',')
        condition_list = request.form.getlist('noOfConditions[]',None)
        no_of_condition = []
        cond_flag = False;
        for i in condition_list:
            no_of_condition = i.split(',')
        operator_dict = {'Equal': ' = ', 'Not Equal': ' <> ', 'Greater than': ' > ', 'Less than': ' < ', 'Greater than or equal': ' >= ', 'Less than or equal': ' <= ', 'Starts With': ' LIKE ', 'Ends With': ' LIKE ', 'Between': ' BETWEEN '}
        cond_query = ''
        
        if data == 'parts':
            cust_list = request.form.get('PartList1',None) 
            if cust_list != 'select file name':
                part_list = oracle_connect.get_part_list(file_name=cust_list, user_name=user_name)
                part_list = "','".join(map(str,part_list))
                part_list = "WHERE PART_NO IN ('"+part_list+"')"

            searched_column_list = [config.PARTS_FIELD.get(i,i) for i in search_to_list]
            searched_column = ','.join(searched_column_list)
            
            if searched_column:
                if part_list:
                    query = "Select TOP 1000 " + searched_column +  " from MMBI_VIEWS_ADCUR.PN_HEADER " + part_list
                else:
                    query = "Select TOP 1000 " + searched_column +  " from MMBI_VIEWS_ADCUR.PN_HEADER "
                    
        else:
            operator_val = ''
            file_field = request.form.get('file_field',None)
            fileList = request.form.get('fileList',None)
            if file_field != 'Select File Type':
                if file_field == 'customer':
                    data_list = oracle_connect.get_cust_list(file_name=fileList, user_name=user_name)
                    data_list = "','".join(map(str,data_list))
                    data_list = " CUST IN ('"+data_list+"')"
                else:
                    data_list = oracle_connect.get_part_list(file_name=fileList, user_name=user_name)
                    data_list = "','".join(map(str,data_list))
                    data_list = " PART_NO IN ('"+data_list+"')"
            
            groupByField = request.form.get('groupByField',None)
            aggr_func = request.form.get('aggr_func',None)
            if groupByField != 'select column' and aggr_func != 'select aggregate function':
                grp_flag = True

            searched_column_list = [config.SALES_FIELD.get(i,i) for i in search_to_list]
            searched_column = ','.join(searched_column_list)
            
            if len(no_of_condition) == 1 and no_of_condition[0] != '':
                cond_flag = True;
                field0 = request.form.get('field0')
                operator0 = request.form.get('operator0')
                value1 = request.form.get('value0_op0')
                and_or_cond0 = request.form.get('options0')
                
                if and_or_cond0:
                    cond_query += " " + and_or_cond0 + " ";
                
                if operator0 == 'Starts With':
                    operator_val = operator_dict[operator0] + "'" + value1 + "%" + "'"
                    cond_query += field0 + operator_val
                elif operator0 == 'Ends With':
                    operator_val = operator_dict[operator0] + "'"  + "%" + value1 + "'" 
                    cond_query += field0 + operator_val
                else:   
                    cond_query += field0 + operator_dict[operator0] + "'" + value1 + "'"
                
                if operator0 == "Between":
                        value2 = request.form.get('value1_op0')
                        cond_query += "' and '" + value2 + "'";
                        
            if len(no_of_condition) > 1:
                for i in no_of_condition[1:]:
                    if i:
                        cond_flag = True;
                        field = request.form.get('field' + str(i))
                        operator = request.form.get('operator' + str(i))
                        value1 = request.form.get('value0_op' + str(i))
                        and_or_cond = request.form.get('options' + str(i))
                        
                        if operator == 'Starts With':
                            operator_val = operator_dict[operator] + "'" + value1 + "%" + "'"
                            cond_query += field + operator_val
                        elif operator == 'Ends With':
                            operator_val = operator_dict[operator] + "'"  + "%" + value1 + "'" 
                            cond_query += field + operator_val
                        else:   
                            cond_query += field + operator_dict[operator] + "'" + value1 + "'"
                        
                        if operator == "Between":
                            value2 = request.form.get('value1_op' + str(i))
                            cond_query += "' and '" + value2 + "'"; 
                        
                        if and_or_cond:
                            cond_query += " " + and_or_cond + " ";
                
                field0 = request.form.get('field0')
                operator0 = request.form.get('operator0')
                value1 = request.form.get('value0_op0')
                and_or_cond0 = request.form.get('options0')
                
                if and_or_cond0:
                    cond_query += " " + and_or_cond0 + " ";
                    
                if operator0 == 'Starts With':
                    operator_val = operator_dict[operator0] + "'" + value1 + "%" + "'"
                    cond_query += field0 + operator_val
                elif operator0 == 'Ends With':
                    operator_val = operator_dict[operator0] + "'"  + "%" + value1 + "'" 
                    cond_query += field0 + operator_val
                else:   
                    cond_query += field0 + operator_dict[operator0] + "'" + value1 + "'"
                
                if operator0 == "Between":
                    value2 = request.form.get('value1_op0')
                    cond_query += "' and '" + value2 + "'";
            
            if searched_column:
                if cond_flag == True and data_list and grp_flag == True:
                    query = "Select TOP 1000 " + searched_column + "," + aggr_func + '(' + groupByField + ')' + " from MMBI_VIEWS_ADCUR.PN_SIS WHERE " + cond_query + " AND " + data_list + " GROUP BY " + searched_column
                elif cond_flag == True and data_list:
                    query = "Select TOP 1000 " + searched_column +  " from MMBI_VIEWS_ADCUR.PN_SIS WHERE " + cond_query + " AND " + data_list
                elif cond_flag == True and grp_flag == True:
                    query = "Select TOP 1000 " + searched_column + "," + aggr_func + '(' + groupByField + ')' + " from MMBI_VIEWS_ADCUR.PN_SIS WHERE " + cond_query + " GROUP BY " + searched_column
                elif data_list and grp_flag == True:
                    query = "Select TOP 1000 " + searched_column + "," + aggr_func + '(' + groupByField + ')' + " from MMBI_VIEWS_ADCUR.PN_SIS WHERE " + data_list + " GROUP BY " + searched_column
                elif cond_flag == True:
                    query = "Select TOP 1000 " + searched_column +  " from MMBI_VIEWS_ADCUR.PN_SIS WHERE " + cond_query 
                elif data_list:
                    query = "Select TOP 1000 " + searched_column +  " from MMBI_VIEWS_ADCUR.PN_SIS WHERE " + data_list
                elif grp_flag == True:
                    query = "Select TOP 1000 " + searched_column + "," + aggr_func + '(' + groupByField + ')' + " from MMBI_VIEWS_ADCUR.PN_SIS GROUP BY " + searched_column
                else:
                    query = "Select TOP 1000 " + searched_column +  " from MMBI_VIEWS_ADCUR.PN_SIS "
                
            if grp_flag == True:
                searched_column_list.append(aggr_func + '(' + groupByField + ')')
                
    return query, searched_column_list
    

if __name__ == '__main__':
    app.run(debug=True)