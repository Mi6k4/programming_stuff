import psycopg2

greenplum = psycopg2.connect(host='',
                             database='warehouse',
                             user='zeppelin',
                             password='')
gp_cursor = greenplum.cursor()
print('I am connected to gp')

distinct_views_sql = """select distinct *
           from (select distinct source_schema || '.' || source_table source,
                                 case
                                     when source_schema || '.' || source_table in
                                          (select schemaname || '.' || matviewname from pg_matviews)
                                         then true
                                     else false end as                is_matview
                 from meta.view_dependencies
                 union all
                 select distinct dependent_schema || '.' || dependent_view dep,
                                 case
                                     when dependent_schema || '.' || dependent_view in
                                          (select schemaname || '.' || matviewname from pg_matviews)
                                         then true
                                     else false end as                     is_matview
                 from meta.view_dependencies
                 ) foo
                 where source not like '% %'
                    and source not similar to '%(«|»)%';"""

gp_cursor.execute(distinct_views_sql)
distinct_views = gp_cursor.fetchall()

dependencies_sql = """select distinct source_schema||'.'||source_table source,
                dependent_schema||'.'||dependent_view dep
from meta.view_dependencies
where source_table not like '% %'
    and dependent_view not like '% %'
    and source_table not similar to '%(«|»)%'
    and dependent_view not similar to '%(«|»)%'
union all
select 'datasets.dataset_losses_pure_test' as source,
    'public.dataset_losses_pure_columnar' as dep
    """

gp_cursor.execute(dependencies_sql)
dependencies_from_db = gp_cursor.fetchall()

gp_cursor.close()
greenplum.close()
print('I have closed the connection')

with open(file='all_tables_list.txt', mode='w') as f:
    for line in distinct_views:
        if str(line) not in ("('one_time_data_load_for_merge.выплаты 4', False)",
                             "('one_time_data_load_for_merge.невыплаты 4', False)"):
            f.write(f"{line}\n")
        else:
            print('im in all tables')
            pass
    print('distinct_views written')

with open(file='all_dep_list.txt', mode='w') as f:
    for line in dependencies_from_db:
        if str(line) in ("('robodops.slot_premium_new_materialized', 'bestdoctor_adminka_consumption.slot_view')",
                         "('one_time_data_load_for_merge.невыплаты 4', 'bestdoctor_bestinsure.csat_pay_not')",
                         "('one_time_data_load_for_merge.выплаты 4', 'bestdoctor_bestinsure.csat_pay')"):
            print('im here')
            pass
        else:
            f.write(f"{line}\n")
    print('dependencies_order written')