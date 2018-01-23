## Create records from HGARC edited spreadsheet
## J. Ammerman
## 2018-01-22
##
## load required libraries
try:
    ## this version requires python 2.7
    import urllib2
    from urllib2 import Request, urlopen
    from urllib import urlencode, quote_plus
except ImportError:
    ## run with Python 3
    from urllib.request import Request, urlopen
    from urllib.parse import urlencode, quote_plus
import pandas as pd
import xml.etree.ElementTree as ET
import os
## change the following value to reflect the desired working directory
#os.chdir('/Volumes/jwa_drive2/Google Drive/git/hgarc')

def sort_marc_tags(record):
    '''
    sort_marc_tags receives a marcxml record and sorts the marc tags numerically.
    it returns the sorted marcxml record
    
    ex: sorted_marcxml_record = sort_marc_tags(marc_xml_record)
    '''
    ## create an empty list
    data = []
    ## get a list of the fields in the record
    fields = list(record)
    ## iterate through the fields
    for i in range(0,len(fields)):
        ## we assign the xml field element to the variable 'elem'
        elem = fields[i]
        ## since the leader doesn't sort numerically, we deal with it as a special case
        ## we append a tuple that contains '000' and the elem ('000',elem) to the data list
        if 'leader' in elem.tag: #== 'http://www.loc.gov/MARC21/slim}leader':
            data.append(('000', elem))
        else:
            ## for all other tags, we can use the marc tag as the value to sort on
            ## we assign the value of the 'tag' attribute to the first element of the
            ## tuple. We assign the xml element for the field to the 'elem' variable
            ## and we append the tuple to the data list 
            attrib = elem.attrib
            for k,v in attrib.items():
                if k == 'tag':
                    data.append((v,elem))
    ## here we sort the data list on the first value of the tuple
    data = sorted(data, key=lambda x: x[0])
    ## we create a new record and iterate through the sorted data list to append each field in 
    ## the correct order
    new_rec = ET.Element('record')
    for i in data:
        new_rec.append(i[1])
    ## return the sorted record
    return(new_rec)

def make_field(d,subfields):
    '''
    make_field returns a MARCXML datafield. Two parameters are passed:
    d - The first parameter is a dict containing three key/value combinations:
        'tag' is the numeric tag for the datafield
        'ind1' is the first indicator for the field
        'ind2' is the second indicator for the field
    subfields - The second paramter is a list of dicts containing the subfield data.
        Multiple dicts for subfields can be passed in the list.
        Each dict contains two key/value conbinations:
        'code' is the code attribute of the subfield
        'text' is the value assigned to the subfield
    ex:
        field = {'tag': '555', 'ind1':' ', 'ind2':' '}
        subfields = [{'code':'3','text': 'Inventory'},{'code':'a','text': 'Inventory available in archive; folder level control'},
                    {'code':'b','text': 'available in archive;'},{'code':'c','text': 'folder level control'},
                    {'code':'u','text': 'http://archives.bu.edu/researchers'}]
        _555 = make_field(field,subfields)
    '''
    if len(subfields) > 0:
        f = ET.Element('datafield')
        f.attrib['tag'] = d['tag']
        f.attrib['ind1'] = d['ind1']
        f.attrib['ind2'] = d['ind2']
        for sub in subfields:
            s = ET.Element('subfield')
            s.attrib['code'] = sub['code']
            s.text = sub['text']
            f.append(s)
    elif len(subfields)==0:
        f = ET.Element('controlfield')
        f.attrib['tag'] = d['tag']
        f.attrib['ind1'] = d['ind1']
        f.attrib['ind2'] = d['ind2']
        f.text = d['text']
    else:
        print(len(subfields))
        print(subfields)
        pass
    return(f)

## we create a DataFrame by reading a spreadsheet of edited HGARC records
## The columns in the spreadsheet are:
## CollectionNum,oclcnum,field_type,marc_tag,ind1,ind2,field_value,sub_code,sub_value
## each field/subfield has its own row. The fields/subfields are read and 
## re-assembed to create a new marcxml record

df = pd.read_excel('hgarc_oclc_recs.xlsx')

## set the opening variables
prev_collectionID = ''
prev_marc_tag = ''
record = ''
file_num = 200
## open the first output file
f = open('hgarc-updated-records'+str(file_num)+'.xml','w')
f.write('<?xml version="1.0" encoding="UTF-8"?>')
collection = ET.Element('collection')
for i in df.index:
    collectionID = df.loc[i,'CollectionNum']
    oclcnum = df.loc[i,'oclcnum']
    if collectionID != prev_collectionID: ## we have a new record
        ## we define record as a string to get us started, but 
        ## we don't want to append the string to the collection
        if type(record) != str:
            collection.append(sort_marc_tags(record))
        ## write out the records in collections of 100 records
        if len(collection) == 100:
            f.write(ET.tostring(collection))
            f.close()
            file_num += 1
            f = open('hgarc-updated-records'+str(file_num)+'.xml','w')
            f.write('<?xml version="1.0" encoding="UTF-8"?>')
            collection = ET.Element('collection')
        print(i,collectionID,oclcnum)
        prev_collectionID = collectionID
        record = ET.Element('record')
    try:
        marc_tag = str(int(df.loc[i,'marc_tag']))
    except Exception as e:
        ## the leader does not have a marc_tag, so we account for that
        marc_tag = ''
    ## need to prefix some marc tags with leading zeros
    pad = '000'
    marc_tag = pad[len(marc_tag):] + marc_tag
    ## consider adding the indicators to the marc_tag
    
    if marc_tag != prev_marc_tag: 
        ## we have a new tag so we append the field to record, and reset variables
        try:
            record.append(field)
        except Exception as e:
            pass ## first iteration will not have a valid field
        prev_marc_tag = marc_tag
        subfields = []
    ## case for leader   
    if df.loc[i,'field_type'] == 'leader':
        marc_tag = 'leader'
        field = ET.Element('leader')
        field.text = df.loc[i,'field_value']
    ## case for controlfield
    elif df.loc[i,'field_type'] == 'controlfield':
        field = ET.Element('controlfield')
        field.attrib['tag'] = marc_tag
        field.text = df.loc[i,'field_value']
    ## case for datafields
    else: ## here we deal with datafields 
        ## for most, we can add multiple subfields to the same marc tag
        ## for a few, we need separate marc tags because they contain
        ## non-repeatable subfields.
        ## These will be in a list that we will check. If in the list,
        ## we will attempt to make the field and begin a new list of 
        ## subfields.
        nr_subfields = ['500','501','502','504','505','506','507',
                        '508','510','511','513','514','515','516',
                        '518','520','524','525','526','530','533',
                        '534','535','536','538','540','541','542',
                        '545','550','552','555','556','561','563',
                        '565','567','580','581','583','585','586','588',
                        '600','610','611','630','647','648','650',
                        '651','655','656','657','658','700','710','711',
                        '720','730','740','751','753','758',
                        '800','810','811','830','850','852'
                       ]
        sub_code = df.loc[i,'sub_code']
        ## check to see if the current marc tag is in the nr_subfields list
        if marc_tag in nr_subfields:
            if sub_code == 'a':
                for subfield in subfields:
                    for k,v in subfield.items():
                        if subfield[k] == 'a':
                            #print(marc_tag,v)
                            field = make_field(d,subfields)
                            record.append(field)
                            subfields = []
            
        d = {'tag':marc_tag,'ind1':str(df.loc[i,'ind1']),
            'ind2':str(df.loc[i,'ind2'])}
        subfields.append({'code':sub_code,'text':df.loc[i,'sub_value']})
        field = make_field(d,subfields)
        prev_marc_tag = marc_tag
collection.append(sort_marc_tags(record))      

f.write(ET.tostring(collection))
f.close()

