import psycopg2
from typing import List

gp_conn = 'postgresql://zeppelin:R63v5NspNsSEem@172.21.216.13:5432/warehouse' # подлючение под пользаком zeppelin


class DwsConn:
    def __init__(self, conn_string):
        self.conn_string = conn_string

    def select(self, query: str) -> List[tuple]:
        with psycopg2.connect(self.conn_string) as conn:
            with conn.cursor() as cursor:
                cursor.execute(query)
                return cursor.fetchall()

    def execute(self, query: str) -> None:
        with psycopg2.connect(self.conn_string) as conn:
            with conn.cursor() as cursor:
                cursor.execute(query)
                conn.commit()




tables_list = [
' create table bestdoctor_bb.bb_telemedicine_consultationtemplate_checkups_expense_patient as select * from bestdoctor_bb_external_tables.instancechangerecord_combined_columnar where content_type_id =    661 ; ' ,
' create table bestdoctor_bb.bb_medical_program_clinicprogramservicetreenode as select * from bestdoctor_bb_external_tables.instancechangerecord_combined_columnar where content_type_id =    671 ; ' ,
' create table bestdoctor_bb.bb_registries_technicalregistryreason as select * from bestdoctor_bb_external_tables.instancechangerecord_combined_columnar where content_type_id =    692 ; ' ,
' create table bestdoctor_bb.bb_documents_documentlink as select * from bestdoctor_bb_external_tables.instancechangerecord_combined_columnar where content_type_id =    693 ; ' ,

]

smaller_list= [
'	create table bestdoctor_bb.bb_doctors_doctorcurator as select * from bestdoctor_bb_external_tables.instancechangerecord_combined_columnar where content_type_id =    340 ; ' ,
' create table bestdoctor_bb.bb_programs_programdirectattachmentstomatology as select * from bestdoctor_bb_external_tables.instancechangerecord_combined_columnar where content_type_id =    409 ; ' ,
' create table bestdoctor_bb.bb_approvals_approvement_appeals as select * from bestdoctor_bb_external_tables.instancechangerecord_combined_columnar where content_type_id =    427 ; ' ,
' create table bestdoctor_bb.bb_symptomchecker_resultsettingsquestion as select * from bestdoctor_bb_external_tables.instancechangerecord_combined_columnar where content_type_id =    486 ; ' ,
' create table bestdoctor_bb.bb_chat_conversation as select * from bestdoctor_bb_external_tables.instancechangerecord_combined_columnar where content_type_id =    501 ; ' ,

]

DWH = DwsConn(conn_string=gp_conn)
for query in tables_list:
    print('excuting query')
    print(query)
    DWH.execute(query)

