require_relative 'exporter'

UNIT_QUERY = 'select distinct skupina from obce'

DATA_QUERY = 'select
	o1.region as region,
  kod_from,
	o1.nazev as o_from,
  suma,
  coalesce((select suma from matice mx where mx.kod_from = m.kod_to and m.kod_from = mx.kod_to), 0) as suma2,
  kod_to,
	o2.nazev as o_to,
  o2.region as region_doj,
	t1.tik as tik,
	t1.tki as tki,
	t1.tii as tii,
	t2.tik as tjk,
	t2.tki as tkj,
	(case when o1.region = o2.region then (select sum(suma) from matice mx join obce ox on (mx.kod_from = ox.kod) join obce ox2 on (mx.kod_to = ox2.kod) where ox.region = o1.region and ox2.region = o1.region)
	else (select sum(suma) from matice mx join obce ox on (mx.kod_from = ox.kod) join obce ox2 on (mx.kod_to = ox2.kod) where ox.skupina = o1.skupina and ox2.skupina = o1.skupina) end) as sum_tij,
	(case when o1.region = o2.region then (select sum(tii) from tiky tx join obce ox on (tx.kod = ox.kod) where ox.region = o1.region)
	else (select sum(tii) from tiky tx join obce ox on (tx.kod = ox.kod) where ox.skupina = o1.skupina) end) as sum_tii
from
	matice m join tiky t1 on (m.kod_from = t1.kod)
	join tiky t2 on (m.kod_to = t2.kod)
	join obce o1 on (m.kod_from = o1.kod)
	join obce o2 on (m.kod_to = o2.kod)
where
	o1.jadro = \'True\' and o2.jadro = \'True\' and o1.skupina=o2.skupina and o1.skupina=?
  order by o1.region, o1.nazev, o2.nazev'

HEADERS = ['Region', 'KODOB_VYJ', 'NAZOB_VYJ', 'SumOfVYJ_DENNE', 'SumOfDOJ_DENNE', 'KODOB_DOJ', 'NAZOB_DOJ', 'REGION_DOJ', 'Tik', 'Tki', 'Tii', 'Tjk', 'Tkj', 'SumOfTij', 'SumOfTii', 'poměr Tij(sumaTij+sumaTii)', 'poměr Tji/(sumaTij+sumaTii)', 'SMART_Tij', 'SMART_Tji', 'INTRAMAX_Tij', 'INTRA_Tji', 'CURDS_Tij', 'CURDS_Tji']
FORMULAS = ['=D$i/(N$i+O$i)', '=E$i/(N$i+O$i)', '=(D$i*D$i)/(M$i*I$i)+(E$i*E$i)/(L$i*J$i)', '=(E$i*E$i)/(M$i*I$i)+(D$i*D$i)/(L$i*J$i)', '=D$i/(M$i*I$i)+E$i/(L$i*J$i)', '=E$i/(M$i*I$i)+D$i/(L$i*J$i)', '=D$i/M$i+D$i/I$i+E$i/L$i+E$i/J$i', '=E$i/M$i+E$i/I$i+D$i/L$i+D$i/J$i']

exporter = Exporter.new 'jadra_sum.xlsx', 'jadra', HEADERS, FORMULAS
exporter.run 'database.sqlite', UNIT_QUERY, DATA_QUERY