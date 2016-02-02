# USAGE:
# ruby ms_autoformat.rb test2/ VOLUMES.xlsx CONTROL.xlsx "creatinine"
# ruby ms_autoformat.rb test2/ VOLUMES.xlsx # if no indicator

require 'rubygems'
require 'csv'
require 'rubyXL'
require 'axlsx'

mydir = ARGV[0]
vols_file = ARGV[1]
set_indicator = false
if ARGV[2] && ARGV[3]
	indicator_file = ARGV[2]
	indicator  = ARGV[3]
	set_indicator = true
end


dict_list = Hash.new { |h,k| h[k] = [] }
raw_files = Hash.new { |h,k| h[k] = [] }
compounds = {}
Dir["#{mydir}*.txt"].each do |ifile|
	filename = nil
	# system 'dos2unix #{ifile}'
	txt_rows = File.readlines(ifile, {:col_sep => "\t"})
	CSV.open("#{ifile}.csv", 'wb') do |csv|
		filename = ifile.split("/")[-1].split(".txt")[0]
		txt_rows.each do |row|
			cols = row.split("\t")
			cols[-1] = cols[-1].split("\r\n")[0]
			raw_files[filename] << cols
			cols.to_csv
			csv << cols
		end
	end

	in_even_compound = false
	compound_num = nil
	compound_name = nil
	fields = {}
	CSV.foreach("#{ifile}.csv") do |row|

		if !row[0].nil? && row[0].start_with?("Compound")
			compound_num = row[0].split(" ")[1].split(":")[0]
			compound_name = row[0].split(":")[1]
			if compound_num.to_i.even?
				compounds[compound_name] = compound_num
				in_even_compound = true
				# puts "COMPOUND #{compound_num}"
				next
			else
				in_even_compound = false
				next
			end
		end
		
		if in_even_compound == true && !row.empty?
			if row[1] && row[1] == "Name"
				for i in 1..row.length-1
					if row[i] == "Name"
						fields["Name"] = i
					elsif row[i] == "Area"
						fields["Area"] = i
					elsif row[i] == "IS Area"
						fields["IS Area"] = i
					elsif row[i] == "Std. Conc"
						fields["Std. Conc"] = i
					end
				end
				next
			elsif (row[0] != "") && (row[1] != "")
				dict_list[compound_name] << [row[fields["Name"]], row[fields["Area"]], row[fields["IS Area"]], row[fields["Std. Conc"]]]
			end
		elsif in_even_compound == false
			next
		end

	end
end

# dict_list.each_key do |key|
# 	puts "#{key} = { #{dict_list[key].join(",")} }"
# 	# 15Oct05_KT_18_PGs6k_MU_124 = {   PGE-M MO,25,77.478,97917.891,  PGD-M MO,25,237.123,79432.523,  d6k-PGF1a MO,5,719.522,91150.852,  dTx MO,5,41355.34,169043.25,  6k-PGF1a MO,5,587.895,45864.465 }
# end


# read the volumes list (format: ID	Sample	Volume(ml)	Treatment	Analytes)
# eg: 1	220-4 0.1	none	DM,EM,IM,Tx,6k
workbook = RubyXL::Parser.parse(vols_file)
volumes = {}
sample_names = {}
treatments = {}
worksheet1 = workbook[0]
worksheet1.each do |row|
	if row && row.cells
		if row[0] && (row[0].value != "ID*")
			if !row[2].value
				next
			else
				id = row[0].value.to_s
				key_id = id.gsub(/[^\d]/,'')
				volumes[key_id] = row[2].value
				sample_names[key_id] = row[1].value
				treatments[key_id] = row[3].value if row[3]
			end
		end
	end
end

# volumes.each_key do |key|
# 	puts "#{key} = #{volumes[key]}" # 124 = 0.1
# end

indicator_volumes = Hash.new { |h,k| h[k] = [] }
indicator_samples = Hash.new { |h,k| h[k] = [] }
if set_indicator == true
	# read the indicator list (format: ID - Filename	- Area - ISTD Area - Area Ratio - Sample ID - Treatment	- Sample Volume (mL) - Spike (ng) - ng/ml	- mg/dl)
	# eg: 001 15Oct05_Creat_KT_Study18_MU_001	4,146,988	6,805,456	0.609	220-4 urine 1	none	0.010	2,500	152340.55	15.23 
	# keep cols: 0,4,9,5
	workbook = RubyXL::Parser.parse(indicator_file)
	worksheet = workbook[0]
	in_data = false
	worksheet.each do |row|
		if row && row.cells
			if row[0] && (row[0].value == "ID*")
				in_data = true
				next
			end

			if in_data == true && row[0] #&& row[9].value
				values = {}
				for i in 1..8
					if row[i].nil?
						value = ""
					else
						value = row[i].value
					end
					values["row#{i}"] = value
				end	
				if (!values["row4"].nil? && !values["row8"].nil? && !values["row7"].nil?)
					ng_ml = values["row4"] * values["row8"] / values["row7"]
					mg_dl = (ng_ml * 100 / 1000000)
					indicator_volumes[row[0].value] = [row[0].value, mg_dl.round(2), values["row1"]]
				else
					indicator_volumes[row[0].value] = [row[0].value, values["row7"], values["row1"]]
				end
				indicator_samples[row[0].value] = row[0].value
			end
		end
	end
	# indicator_volumes.each_key do |key|
	# 	puts "#{key} = #{indicator_volumes[key].join(",")}"
	# 	# 15Oct05_Creat_KT_Study18_MU_124 = 231-3 urine 4,9.2568380893374, High-salt+celecoxib 
	# end
	# indicator_samples.each_key do |key|
	# 	puts "#{key} = #{indicator_samples[key]}"
	# 	# 124 = 15Oct05_Creat_KT_Study18_MU_124
	# end
end


# output
dir_name = mydir.split("/")[0]
ofile = "#{mydir}/#{dir_name}_DATA.xlsx"
results_xlsx = Axlsx::Package.new
results_wb = results_xlsx.workbook
title = results_wb.styles.add_style(:b => true, :alignment=>{:horizontal => :center})
data = results_wb.styles.add_style(:alignment=>{:horizontal => :left})

# precreate the summary and final sheets, to be the first two sheets
if set_indicator == true
	final_sheet = results_wb.add_worksheet(:name => "Final")
end
summary_sheet = results_wb.add_worksheet(:name => "Summary")

# create compound sheets
vols_list = Hash.new { |h,k| h[k] = [] }
sample_list = Hash.new { |h,k| h[k] = [] }
fnames_by_sample = {}
dict_list.each_key do |compound|
	results_wb.add_worksheet(:name => compound) do |sheet|
		sheet.add_row ["Sample #","File Name","Spike (ng)","Vol (ml)","Endogenous Area","Spike Area","Endog/Spike-RC","ng/ml"], :style => title

		bias = nil
		bias_set = false
		
		dict_list[compound].each do |sample_arr|
			vol = ""
			fname = sample_arr[0]
			sample = fname.split("_")[-1]
			if !sample.include? "RC"
				if sample =~ /\d/
					sample = sample.gsub(/[^\d]/,'')
					# p "HERE #{fname} >> #{sample}"
				else
					# p "BEFORE #{fname} >> #{sample}"
					sample = fname.split("_")[-2].sub(/[^\d]/,'')
					# p "THERE #{fname} >> #{sample}"
				end
			end
			
			if sample.start_with? "0"
				sample = sample.sub(/^[0]*/,'')
			end

			spike = sample_arr[3].to_f

			if volumes.has_key?(sample)
				vol = volumes[sample].to_f
			else
				vol = ""
			end

			if bias_set == false && (sample.include? "RC")
				bias_set = true
				if sample_arr[1]!="" && sample_arr[2]!=""
					e_area = sample_arr[1].to_f
					s_area = sample_arr[2].to_f
					bias = e_area/s_area
				else
					s_area = sample_arr[2]
					e_area = ""
					bias = 0
				end
				endo_spike_ratio = bias.round(5)
				f_vol = ""
				sheet.add_row [sample, fname, spike, vol, e_area, s_area, endo_spike_ratio, f_vol], :style => data
				next
			end
			
			if bias_set == true && (!sample.include? "RC")
				if sample_arr[1]!="" && sample_arr[2]!=""
					e_area = sample_arr[1].to_f
					s_area = sample_arr[2].to_f
					endo_spike_ratio = (e_area/s_area)-bias
					if vol != "" && vol && !vol.nil?
						f_vol = spike * endo_spike_ratio / vol
					end
					p sample if f_vol.nil?
				else
					s_area = sample_arr[2]
					e_area = ""
					endo_spike_ratio = 0
					f_vol = 0
				end
				sheet.add_row [sample, fname, spike, vol, e_area, s_area, endo_spike_ratio.round(5), f_vol.round(2)], :style => data
				vols_list[sample] << f_vol
				sample_list[sample] << sample
				fnames_by_sample[sample] = fname
			end
		end
	end
end

#create raw compound sheets
raw_files.each_key do |raw_file|
	if raw_file.size > 30
		raw_file_tmp = raw_file.split(//).last(30).join("").to_s
	else
		raw_file_tmp = raw_file
	end
	results_wb.add_worksheet(:name => raw_file_tmp) do |sheet|	
		raw_files[raw_file].each do |row|
			sheet.add_row row
		end
	end
end

# create summary table with all compound vols
stop1 = 2+compounds.keys.size-1 #6
start2 = stop1 + 2 #8
stop2 = start2 + compounds.keys.size - 1 #12
new_vols_with_indic = Hash.new { |h,k| h[k] = [] }

# summary_sheet = results_wb.add_worksheet(:name => "Summary")
if set_indicator == true
	summary_sheet.add_row ["","",(0..compounds.keys.size-1).map{|i| "ng/ml"},"",(0..compounds.keys.size-1).map{|i| "ng/mg #{indicator}"}].flatten, :style => title
	summary_sheet.merge_cells summary_sheet.rows.first.cells[(2..stop1)]
	summary_sheet.merge_cells summary_sheet.rows.first.cells[(start2..stop2)]
	summary_sheet.add_row ["Sample #","ID",dict_list.keys,"#{indicator}(mg/dl)",dict_list.keys].flatten, :style => title, :widths => [:auto]
else
	summary_sheet.add_row ["","",(0..compounds.keys.size-1).map{|i| "ng/ml"}].flatten, :style => title, :widths => [:auto]
	summary_sheet.merge_cells summary_sheet.rows.first.cells[(2..stop1)]
	summary_sheet.add_row ["Sample #","ID",dict_list.keys].flatten, :style => title
end

vols_list.each_key do |sample|
	# fname_id = sample.split("_")[-1].gsub(/[^0-9]/i, '').sub(/^[0]*/,'').to_i
	if set_indicator == true && indicator_samples.has_key?(sample.to_i) 
		indicator_vol = indicator_volumes[sample.to_i][1]
		summary_sheet.add_row [sample_list[sample][0], fnames_by_sample[sample.to_s], vols_list[sample].map { |i| i.round(2) }, indicator_vol.round(2), vols_list[sample].map { |i| (i*100/indicator_vol).round(2) }].flatten, :style => data
		vols_list[sample].each do |i|
			if (i*100/indicator_vol) < 0
				final_vol = "BLQ"
			else
				final_vol = (i*100/indicator_vol).round(2)
			end
			new_vols_with_indic[sample_list[sample][0]] << final_vol
		end
	else
		summary_sheet.add_row [sample_list[sample][0], fnames_by_sample[sample.to_s], vols_list[sample].map { |i| i.round(2) }].flatten, :style => data
	end
end


# create final table if indicator is set
if set_indicator == true
	# final_sheet = results_wb.add_worksheet(:name => "Final")
	final_sheet.add_row ["","","",(0..compounds.keys.size-1).map{|i| "ng/ml #{indicator}"}].flatten, :style => title
	final_sheet.merge_cells final_sheet.rows.first.cells[(3..3+compounds.keys.size-1)]
	final_sheet.add_row ["Sample #","ID","Treatment",dict_list.keys].flatten, :style => title, :widths => [:auto]

	new_vols_with_indic.each_key do |sample|
		# sample_3pos = sprintf '%03d', sample
				
		final_sheet.add_row [sample, sample_names[sample], treatments[sample], new_vols_with_indic[sample].map { |i| i }].flatten, :style => data
	end
end

# write xlsx file
results_xlsx.serialize(ofile)

