require_relative '../lib/exporter'

UNIT_QUERY = 'select distinct skupina from obce'

DATA_QUERY = 'select
	o1.region as region,
	(select count(*) from obce obc where obc.jadro = \'True\' and obc.region=o1.region) as pocet_jader,
	(case when
			(o1.region < o2.region) or
			(o1.region = o2.region and o1.nazev < o2.nazev)
		then 1
		else 2 end) as smer,
  kod_from,
	o1.nazev as o_from,
	o1.pos_x as from_x,
	o1.pos_y as from_y,
  suma,
  coalesce((select suma from matice mx where mx.kod_from = m.kod_to and m.kod_from = mx.kod_to), 0) as suma2,
  kod_to,
	o2.nazev as o_to,
	o2.pos_x as from_x,
	o2.pos_y as from_y,
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

def x(sym)
	HEADERS.each do |h|
		if h[0].eql? sym
			return h[2]
		end
	end
end

HEADERS = [
		[:region, 'Region'],
		[:pocet_jader, 'Pocet_jader'],
		[:smer, 'Smer'],
		[:kod_obce_z, 'KODOB_VYJ'],
		[:nazev_obce_z, 'NAZOB_VYJ'],
		[:obec_z_x, 'FROM_X'],
		[:obec_z_y, 'FROM_Y'],
		[:sum_vyj, 'SumOfVYJ_DENNE'],
		[:sum_doj, 'SumOfDOJ_DENNE'],
		[:kod_obce_do, 'KODOB_DOJ'],
		[:nazev_obce_do, 'NAZOB_DOJ'],
		[:obec_do_x, 'TO_X'],
		[:obec_do_y, 'TO_Y'],
		[:region_do, 'REGION_DOJ'],
		[:tik, 'Tik'],
		[:tki, 'Tki'],
		[:tii, 'Tii'],
		[:tjk, 'Tjk'],
		[:tkj, 'Tkj'],
		[:sum_tij, 'SumOfTij'],
		[:sum_tii, 'SumOfTii'],
		# formulas
		[:pomer_tij, 'poměr Tij(sumaTij+sumaTii)'],
		[:pomer_tji, 'poměr Tji/(sumaTij+sumaTii)'],
		[:smart_tij, 'SMART_Tij'],
		[:smart_tji, 'SMART_Tji'],
		[:intramax_tij, 'INTRAMAX_Tij'],
		[:intramax_tji, 'INTRA_Tji'],
		[:curds_tij, 'CURDS_Tij'],
		[:curds_tji, 'CURDS_Tji'],
		[:smart_tij_n, 'SMART_Tij_Norm'],
		[:intramax_tij_n, 'INTRAMAX_Tij_Norm'],
		[:curds_tij_n, 'CURDS_Tij_Norm'],
		[:kategorie_regionu, 'Kategorie_regionu']
]

HEADERS.each_index do |idx|
	a = (idx.div 26) - 1
	b = idx % 26
	char = '' if a < 0
	char = ('A'.ord + a).chr if a >= 0
	char << ('A'.ord + b).chr
	HEADERS[idx] << char
end


FORMULAS = [
		"=#{x :sum_vyj}$i/(#{x :sum_tij}$i+#{x :sum_tii}$i)",
		"=#{x :sum_doj}$i/(#{x :sum_tij}$i+#{x :sum_tii}$i)",
		"=(#{x :sum_vyj}$i*#{x :sum_vyj}$i)/(#{x :tkj}$i*#{x :tik}$i)+(#{x :sum_doj}$i*#{x :sum_doj}$i)/(#{x :tjk}$i*#{x :tki}$i)",
		"=(#{x :sum_doj}$i*#{x :sum_doj}$i)/(#{x :tkj}$i*#{x :tik}$i)+(#{x :sum_vyj}$i*#{x :sum_vyj}$i)/(#{x :tjk}$i*#{x :tki}$i)",
		"=#{x :sum_vyj}$i/(#{x :tkj}$i*#{x :tik}$i)+#{x :sum_doj}$i/(#{x :tjk}$i*#{x :tki}$i)",
		"=#{x :sum_doj}$i/(#{x :tkj}$i*#{x :tik}$i)+#{x :sum_vyj}$i/(#{x :tjk}$i*#{x :tki}$i)",
		"=#{x :sum_vyj}$i/#{x :tkj}$i+#{x :sum_vyj}$i/#{x :tik}$i+#{x :sum_doj}$i/#{x :tjk}$i+#{x :sum_doj}$i/#{x :tki}$i",
		"=#{x :sum_doj}$i/#{x :tkj}$i+#{x :sum_doj}$i/#{x :tik}$i+#{x :sum_vyj}$i/#{x :tjk}$i+#{x :sum_vyj}$i/#{x :tki}$i",
		"=(#{x :smart_tij}$i/MAX(#{x :smart_tij}1:#{x :smart_tij}10000))*100",
		"=(#{x :intramax_tij}$i/MAX(#{x :intramax_tij}1:#{x :intramax_tij}10000))*100",
		"=(#{x :curds_tij}$i/MAX(#{x :curds_tij}1:#{x :curds_tij}10000))*100",
		"=(min(#{x :pocet_jader}$i, 3))"
]

exporter = Exporter.new 'jadra_sum.xlsx', 'jadra', HEADERS.map {|h| h[1]}, FORMULAS
exporter.run 'database.sqlite', UNIT_QUERY, DATA_QUERY