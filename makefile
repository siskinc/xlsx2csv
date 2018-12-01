build:
	g++ -std=c++14 -lxlnt main.cpp -o xlsx2csv
install:
	sudo cp xlsx2csv /usr/local/bin/xlsx2csv