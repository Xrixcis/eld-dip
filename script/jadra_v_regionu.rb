require_relative '../lib/exporter'

UNIT_QUERY = 'select distinct region from obce'

DATA_QUERY = 'select region, count(*) as jader
from
	obce o
where
	o.region = ? and o.jadro = \'True\'
  order by o.region'

HEADERS = ['Region', 'Jader', 'Kategorie']
FORMULAS = ['=MIN(B$i; 3)']

exporter = Exporter.new 'jadra_v_regionu_sum.xlsx', 'jadra_v_regionu', HEADERS, FORMULAS
exporter.run 'database.sqlite', UNIT_QUERY, DATA_QUERY