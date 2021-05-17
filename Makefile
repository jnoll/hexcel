build:
	stack build

test-hexcel: build
	stack exec hexcel -- --sheet Sheet1 --input rubric.xlsx --output example.xlsx --values values.csv

test-hexcel2: build
	cp rubric.xlsx input.xlsx
	stack exec hexcel -- --sheet Sheet1 --input input.xlsx --output input.xlsx --values values.csv

test-xlsx: build
	echo "" | stack exec rubric2xlsx -- --rubric rubric.csv  --output rubric.xlsx 

test-rupdate: build
	cp rubric.xlsx input.xlsx
	stack exec rupdate -- --sheet Sheet1 --input input.xlsx --output output.xlsx --values marks.csv

test-rupdate2: build
	cp rubric.xlsx input.xlsx
	stack exec rupdate -- --sheet Sheet1 --input input.xlsx --output input.xlsx --values marks.csv
