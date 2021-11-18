build:
	stack build

# Populate from (row, col, val) triples
test-hexcel: build
	stack exec hexcel -- --sheet Sheet1 --input rubric.xlsx --output example.xlsx --values values.csv

# Populate from (row, col, val) triples, input=output
test-hexcel2: build
	cp rubric.xlsx input.xlsx
	stack exec hexcel -- --sheet Sheet1 --input input.xlsx --output input.xlsx --values values.csv

# Populate from matching csv
test-hexcel3: build
	cp rubric.xlsx input.xlsx
	stack exec hexcel -- --sheet Sheet1 --input input.xlsx --update --output output.xlsx --values rubric-vals.csv

# Create rubric from matching csv
test-xlsx: build
	echo "" | stack exec rubric2xlsx -- --rubric rubric.csv  --output rubric.xlsx 

# Populate rubric from (section, question, mark, comment) tuples.
test-rupdate: build
	cp rubric.xlsx input.xlsx
	stack exec rupdate -- --sheet Sheet1 --input input.xlsx --output output.xlsx --values marks.csv

# Populate rubric from (section, question, mark, comment) tuples; input=output.
test-rupdate2: build
	cp rubric.xlsx input.xlsx
	stack exec rupdate -- --sheet Sheet1 --input input.xlsx --output input.xlsx --values marks.csv

# Populate rubric from yaml, input=output
test-yupdate: build
	cp rubric.xlsx input.xlsx
	stack exec rupdate -- --sheet Sheet1 --input input.xlsx --output input.xlsx --values example.yml

# Populate rubric from yaml on stdin, input=output
test-yupdate2: build
	cp rubric.xlsx input.xlsx
	cat example.yml | stack exec rupdate -- --sheet Sheet1 --input input.xlsx --output input.xlsx --yaml --values - 

# 
test-update_marks: build
	stack exec update_marks -- --sheet Sheet1 --input marksheet.xlsx --output output.xlsx --values scores.csv

test-update_marks2: build
	cat scores.csv | stack exec update_marks -- --sheet Sheet1 --input marksheet.xlsx --output output.xlsx --values -
