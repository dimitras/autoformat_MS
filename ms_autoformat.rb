# USAGE:
# ruby ms_autoformat.rb input/enrichr_GO_tables_for_filtered_genes.xlsx output/filtered_GO_tables_for_filtered2.xlsx "creatinine"

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
	# puts ifile
	filename = nil
	txt_rows = File.readlines(ifile, {:col_sep => "\t"})
	CSV.open("#{ifile}.csv", 'wb') do |csv|
		filename = ifile.split("/")[-1].split(".txt")[0]
		txt_rows.each do |row|
			cols = row.split("\t")
			raw_files[filename] << cols
			cols.to_csv
			csv << cols
		end
	end

	in_even_compound = false
	compound_num = nil
	compound_name = nil
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

		if in_even_compound == true 
			if row[2] && !dict_list.has_key?(row[2])
				if row[2] == "Name"
					next
				elsif row[0] != ""
					# puts row.inspect
					# dict_list[row[2]] << [compound_name, row[4], row[6..7]]
					dict_list[compound_name] << [row[2], row[4], row[6..7]]
				end
			elsif row[2] && dict_list.has_key?(row[2])
				if row[2] == "Name"
					next
				elsif row[0] != ""
					# puts row.inspect
					dict_list[compound_name] << [row[2], row[4], row[6..7]]
				end
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


# read the volume lists (format: IT MaT ID	Sample	Treatment	Analytes 	volumes(ml)	IT MaT ID	Sample	Treatment	Analytes 	volumes(ml))
workbook = RubyXL::Parser.parse(vols_file)
volumes = {}
worksheet1 = workbook[0]
worksheet1.each do |row|
	if row && row.cells
		if row[0] && (row[0].value != "IT MaT ID")
			volumes[row[0].value] = row[4].value if row[4]
			volumes[row[5].value] = row[9].value if row[5] && row[9]
		end
	end
end

# volumes.each_key do |key|
# 	puts "#{key} = #{volumes[key]}" # 124 = 0.1
# end

indicator_volumes = Hash.new { |h,k| h[k] = [] }
indicator_samples = Hash.new { |h,k| h[k] = [] }
if set_indicator == true
	# read the indicator list (format: Filename	Area	ISTD Area	Area Ratio	Sample ID	Treatment	Sample Volume (mL)	Spike (ng)	ng/ml	mg/dl ) 
	# keep cols: 0,4,9
	workbook = RubyXL::Parser.parse(indicator_file)
	worksheet = workbook[0]
	in_data = false
	worksheet.each do |row|
		if row && row.cells
			if row[0] && (row[0].value == "Filename")
				in_data = true
				next
			elsif row[4] && (row[4].value == "Sample ID")
				in_data = true
				next
			end
			if in_data == true && row[0]
				indicator_volumes[row[0].value] = [row[4].value, row[9].value, row[5].value] if row[4] && row[9] && row[5]
				indicator_samples[row[0].value.split("_")[-1].gsub(/[^0-9]/i, '')] = row[0].value
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
ofile = "results.xlsx"
results_xlsx = Axlsx::Package.new
results_wb = results_xlsx.workbook
title = results_wb.styles.add_style(:b => true, :alignment=>{:horizontal => :center})
data = results_wb.styles.add_style(:alignment=>{:horizontal => :left})

# create compound sheets
raw_files.each_key do |raw_file|
	results_wb.add_worksheet(:name => "Raw_#{raw_file}") do |sheet|
		raw_files[raw_file].each do |row|
			sheet.add_row row
		end
	end
end


vols_list = Hash.new { |h,k| h[k] = [] }
sample_list = Hash.new { |h,k| h[k] = [] }
dict_list.each_key do |compound|
	results_wb.add_worksheet(:name => compound) do |sheet|
		sheet.add_row ["Sample #","File Name","Spike (ng)","Vol (ml)","Endogenous Area","Spike Area","Endog/Spike","ng/ml"], :style => title

		bias = nil
		bias_set = false
		
		dict_list[compound].each do |sample_arr|
			vol = ""
			fname = sample_arr[0]
			sample = fname.split("_")[-1].to_i
			spike = sample_arr[1].to_f
			if volumes.has_key?(sample)
				vol = volumes[sample].to_f
			else
				vol = ""
			end

			if bias_set == false && sample == 0 #RC
				bias_set = true
				if sample_arr[2][0]!="" && sample_arr[2][1]!=""
					e_area = sample_arr[2][0].to_f
					s_area = sample_arr[2][1].to_f
					bias = (e_area/s_area)
				else
					s_area = sample_arr[2][1]
					e_area = ""
					bias = 0
				end
				endo_spike_ratio = bias
				f_vol = ""
				# puts vol
				sheet.add_row [sample, fname, spike, vol, e_area, s_area, endo_spike_ratio.round(5), f_vol], :style => data
				next
			end
# puts vol			
			if bias_set == true && sample != "RC"
				# puts vol
				if sample_arr[2][0]!="" && sample_arr[2][1]!=""
					e_area = sample_arr[2][0].to_f
					s_area = sample_arr[2][1].to_f
					endo_spike_ratio = ((e_area/s_area)-bias)
					if vol != "" && vol && !vol.nil?
						# puts "IN"
						f_vol = ((spike * endo_spike_ratio) / vol)
						# puts f_vol
					end

				else
					s_area = sample_arr[2][1]
					e_area = ""
					endo_spike_ratio = 0
					f_vol = 0
				end
				# puts f_vol
				sheet.add_row [sample, fname, spike, vol, e_area, s_area, endo_spike_ratio.round(5), f_vol.round(2)], :style => data
				vols_list[fname] << f_vol
				sample_list[fname] << sample
			end
		end
	end
end


# create summary table with all compound vols
stop1 = 2+compounds.keys.size-1 #6
start2 = stop1 + 2 #8
stop2 = start2 + compounds.keys.size - 1 #12
new_vols_with_indic = Hash.new { |h,k| h[k] = [] }

summary_sheet = results_wb.add_worksheet(:name => "Summary")
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

vols_list.each_key do |fname|
	fname_id = fname.split("_")[-1].gsub(/[^0-9]/i, '')
	if set_indicator == true && indicator_samples.has_key?(fname_id) 
		indicator_vol = indicator_volumes[indicator_samples[fname_id]][1]
		summary_sheet.add_row [sample_list[fname][0], fname, vols_list[fname].map { |i| i.round(2) }, indicator_vol.round(2), vols_list[fname].map { |i| (i*100/indicator_vol).round(2) }].flatten, :style => data
		vols_list[fname].each do |i|
			if (i*100/indicator_vol) < 0
				final_vol = "BLQ"
			else
				final_vol = (i*100/indicator_vol).round(2)
			end
			new_vols_with_indic[sample_list[fname][0]] << final_vol
		end
	else
		summary_sheet.add_row [sample_list[fname][0], fname, vols_list[fname].map { |i| i.round(2) }].flatten, :style => data
	end
end


# create final table if indicator is set
if set_indicator == true
	final_sheet = results_wb.add_worksheet(:name => "Final")
	final_sheet.add_row ["","","",(0..compounds.keys.size-1).map{|i| "ng/ml #{indicator}"}].flatten, :style => title
	final_sheet.merge_cells final_sheet.rows.first.cells[(3..3+compounds.keys.size-1)]
	final_sheet.add_row ["Sample","ID","Treatment",dict_list.keys].flatten, :style => title, :widths => [:auto]

	new_vols_with_indic.each_key do |sample|
		sample_3pos = sprintf '%03d', sample
		# if indicator_samples.has_key?(sample) 
			final_sheet.add_row [sample, indicator_volumes[indicator_samples[sample_3pos]][0], indicator_volumes[indicator_samples[sample_3pos]][2], new_vols_with_indic[sample].map { |i| i }].flatten, :style => data
		# end
	end
end


# write xlsx file
results_xlsx.serialize(ofile)

