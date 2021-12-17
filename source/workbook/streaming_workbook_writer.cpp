// Copyright (c) 2017-2021 Thomas Fussell
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#include <fstream>

#include <xlnt/cell/cell.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/utils/optional.hpp>
#include <xlnt/workbook/streaming_workbook_writer.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <detail/implementations/cell_impl.hpp>
#include <detail/implementations/workbook_impl.hpp>
#include <detail/implementations/worksheet_impl.hpp>
#include <detail/serialization/open_stream.hpp>
#include <detail/serialization/vector_streambuf.hpp>
#include <detail/serialization/xlsx_producer.hpp>

namespace xlnt {

streaming_workbook_writer::streaming_workbook_writer()
{
}

streaming_workbook_writer::~streaming_workbook_writer()
{
    close();
}

void streaming_workbook_writer::close()
{
    if (producer_)
    {
        end_current_streaming();
        producer_->populate_archive(true);
        producer_.reset(nullptr);
        stream_buffer_.reset(nullptr);
    }
}

cell streaming_workbook_writer::add_cell(const cell_reference &ref)
{
    return producer_->add_cell(ref);
}

void streaming_workbook_writer::add_comment(const cell_reference &ref, const std::string &comment)
{
    //start comment stream if this is the first comment (also ends sheet streaming in producer)
    if (!streaming_comments_)
    {
        producer_->stream_comments_begin();
        streaming_sheet_ = false;
        streaming_comments_ = true;
    }

    producer_->stream_comment(ref, comment);
}

worksheet streaming_workbook_writer::add_worksheet(const std::string &title)
{
    end_current_streaming();

    auto sheet = producer_->add_worksheet(title);
    workbook_->register_worksheet(sheet);
    producer_->stream_worksheet_begin();
    streaming_sheet_ = true;
    return sheet;
}

void streaming_workbook_writer::open(std::vector<std::uint8_t> &data)
{
    stream_buffer_.reset(new detail::vector_ostreambuf(data));
    stream_.reset(new std::ostream(stream_buffer_.get()));
    open(*stream_);
}

void streaming_workbook_writer::open(const std::string &filename)
{
    stream_.reset(new std::ofstream());
    xlnt::detail::open_stream(static_cast<std::ofstream &>(*stream_), filename);
    open(*stream_);
}

#ifdef _MSC_VER
void streaming_workbook_writer::open(const std::wstring &filename)
{
    stream_.reset(new std::ofstream());
    xlnt::detail::open_stream(static_cast<std::ofstream &>(*stream_), filename);
    open(*stream_);
}
#endif

void streaming_workbook_writer::open(const xlnt::path &filename)
{
    stream_.reset(new std::ofstream());
    xlnt::detail::open_stream(static_cast<std::ofstream &>(*stream_), filename.string());
    open(*stream_);
}

void streaming_workbook_writer::open(std::ostream &stream)
{
    workbook_.reset(new workbook());
    producer_.reset(new detail::xlsx_producer(*workbook_));
    producer_->open(stream);
    producer_->current_worksheet_ = new detail::worksheet_impl(workbook_.get(), 1, "Sheet1");
    producer_->current_cell_ = new detail::cell_impl();
    producer_->current_cell_->parent_ = producer_->current_worksheet_;
    producer_->current_row_ = 0;
    workbook_->remove_sheet(workbook_->active_sheet());
}

void streaming_workbook_writer::end_current_streaming()
{
    bool finished = false;

    if (streaming_sheet_)
    {
        producer_->stream_worksheet_end();
        streaming_sheet_ = false;
        finished = true;
    }
    if (streaming_comments_)
    {
        producer_->stream_comments_end();
        streaming_comments_ = false;
        finished = true;
    }

    if (finished)
        producer_->stream_worksheet_rels();
}

} // namespace xlnt
