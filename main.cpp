#include <xlnt/xlnt.hpp>
#include <vector>
#include <iostream>
#include <string>
#include <cstring>
#include <fstream>
#include "cmdline/cmdline.h"

int main(int argc, char **argv)
{
    cmdline::parser cmdline_parser;
    cmdline_parser.add<std::string>("filepath", 'f', "xlsx file path", true);
    cmdline_parser.add<std::string>("sheet_name", 's', "the convert sheet name", false, "");
    cmdline_parser.add<std::size_t>("sheet_index", 'i', "the convert sheet index", false, 0);
    cmdline_parser.add<std::string>("out_file", 'o', "the output file path", false, "output.csv");
    cmdline_parser.parse_check(argc, argv);
    std::string filename = cmdline_parser.get<std::string>("filepath");
    std::string sheet_name = cmdline_parser.get<std::string>("sheet_name");
    std::size_t sheet_index = cmdline_parser.get<std::size_t>("sheet_index");
    std::string output_file = cmdline_parser.get<std::string>("out_file");
    std::ifstream fin(filename);
    if(!fin)
    {
        std::clog << "Please input the xlsx file is exist" << std::endl;
        return -1;
    }
    fin.close();
    auto wb = xlnt::workbook();
    wb.load(filename);
    xlnt::worksheet ws;
    if("" == sheet_name)
    {
        ws = wb.active_sheet();
    }
    else if("" != sheet_name)
    {
        if (!wb.contains(sheet_name))
        {
            std::clog << "Please input a exist sheet name" << std::endl;
            return -1;
        }
        ws = wb.sheet_by_title(sheet_name);
    }
    else if(0 != sheet_index)
    {
        if(wb.sheet_count() >= sheet_index)
        {
            std::clog << "sheet_index is error, more than the number of sheet" << std::endl;
            return -1;
        }
        ws = wb.sheet_by_id(sheet_index);
    }
    int maxlen = 0;
    std::ofstream fout(output_file);
    std::stringstream temp_str;
    for (auto row : ws.rows(false)) 
    {
        temp_str.str("");
        int count = 0;
        for (auto cell : row)
        {
            if(count >= 0)
            {
                count++;
                temp_str << ",";
            }
            temp_str << cell.to_string();
        }
        fout << temp_str.str() << '\n';
    }
    return 0;
}