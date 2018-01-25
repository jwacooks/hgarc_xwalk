## Export MarcXML records to excel spreadsheet
## j. ammerman
## 2018-01-10

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
import glob

## estabish initial variables, lists, datframes, etc.
## file_list builds a list on the pattern 'hgarc_1*.xml'. 
file_list = glob.glob('hgarc_1*.xml')
rec_count = 0
col_num = 0
columns = [u'CollectionNum', u'oclcnum', u'field_type', 
           u'marc_tag', u'ind1', u'ind2',u'field_value', 
           u'sub_code',u'sub_value']
df = pd.DataFrame()
out_file_num = 0
## iterate through each file
for f in file_list[0:]:
    cnum = 0
    collection = ET.parse(f).getroot()
    #rec_count += len(list(collection))
    records = list(collection)

    for record in records[0:]:
        rec_count += 1
        print(rec_count)
        if rec_count == 25:
            ## we write change the column order of the dataframe 
            rec_count = 0
            print('writing to excel')
            df = df[[u'CollectionNum', u'oclcnum', u'field_type', 
                           u'marc_tag', u'ind1', u'ind2',u'field_value', 
                           u'sub_code',u'sub_value']]
            ## and write it to an excel spreadsheet
            df.to_excel('hgarc_oclc_recs_'+str(out_file_num)+'.xlsx')
            df = pd.DataFrame()
            out_file_num += 1
            
        cnum += 1 
        ## we look for the 099 tag where we have stored the collection number
        _099 = record.find('*/[@tag="099"]/*/[@code="a"]')
        try:
            col_num = _099.text
        except Exception as e:
            col_num = ('0000'+str(cnum))[-4:]
            pass
        fields = list(record)
        ## we have three types of fields: leader, controlfield, and datafield
        ## the type is stored in the field.tag. We grab that and build a case 
        ## for each type. For each, we create a dict that we use to append 
        ## to a DataFrame that will eventually be written out to a spreadsheet
        for field in fields:
            d = {}
            field_type = field.tag
            field_type = field_type.replace('{http://worldcat.org/rb}','')
            field_type = field_type.replace('{http://www.loc.gov/MARC21/slim}','')
            
            #col_num = ('0000'+str(cnum))[-4:]
            oclcnum = fields[1].text
            ## leader has no subfields or marc tag number
            if field_type == 'leader':
                d['CollectionNum'] = col_num
                d['oclcnum'] = oclcnum
                d['field_type'] = field_type
                d['marc_tag'] = ''
                d['ind1'] = ''
                d['ind2'] = ''
                d['field_value'] = field.text
                #print(oclcnum,field_type,field.text)
                df = df.append(d,ignore_index=True)
            ## controlfields place the marc tag in the 'tag' attribute
            ## the value is in the field.text
            ## they have no subfields, so we create a single dict
            elif field_type == 'controlfield':
                d['CollectionNum'] = col_num
                d['oclcnum'] = oclcnum
                d['field_type'] = field_type
                d['marc_tag'] = field.attrib['tag']
                d['field_value'] = field.text
                d['sub_code'] = ''
                d['sub_value'] = ''
                d['ind1'] = ''
                d['ind2'] = ''
                #print(oclcnum,field_type,field.attrib['tag'],field.text)
                df = df.append(d,ignore_index=True)
            else:
                ## datafields have subfields. For each subfield, we create
                ## a dict that is added to the dataframe
                subfields = list(field)
                for subfield in subfields:
                    try:
                        d['CollectionNum'] = col_num
                        d['oclcnum'] = oclcnum
                        d['field_type'] = field_type
                        d['marc_tag'] = field.attrib['tag']
                        d['ind1'] = field.attrib['ind1']
                        d['ind2'] = field.attrib['ind2']
                        d['field_value'] = ''
                        d['sub_code'] = subfield.attrib['code']
                        d['sub_value'] = subfield.text
                        #print(d)
                        ## important to append this for each subfield.
                        df = df.append(d,ignore_index=True)
                    except Exception as e:
                        print(e)
                    
## we write change the column order of the dataframe      
df = df[[u'CollectionNum', u'oclcnum', u'field_type', 
               u'marc_tag', u'ind1', u'ind2',u'field_value', 
               u'sub_code',u'sub_value']]
## and write it to an excel spreadsheet
df.to_excel('hgarc_oclc_recs_'+str(out_file_num)+'.xlsx')

