require 'stringio'
require 'spreadsheet/excel/writer/biff8'
require 'spreadsheet/excel/internals/biff8'

module Spreadsheet
  module Excel
    module Writer
##
# Writer class for Excel Worksheets. Most write_* method correspond to an
# Excel-Record/Opcode. You should not need to call any of its methods directly.
# If you think you do, look at #write_worksheet
class Worksheet
  include Biff8
  include Internals
  include Internals::Biff8
  attr_reader :worksheet
  def initialize workbook, worksheet
    @workbook = workbook
    @worksheet = worksheet
    @io = StringIO.new ''
    @biff_version = 0x0600
    @bof = 0x0809
    @build_id = 3515
    @build_year = 1996
    @bof_types = {
      :globals      => 0x0005,
      :visual_basic => 0x0006,
      :worksheet    => 0x0010,
      :chart        => 0x0020,
      :macro_sheet  => 0x0040,
      :workspace    => 0x0100,
    }
  end
  ##
  # The number of bytes needed to write a Boundsheet record for this Worksheet
  # Used by Writer::Worksheet to calculate various offsets.
  def boundsheet_size
    name.size + 10
  end
  def data
    @io.rewind
    @io.read
  end
  def encode_date date
    return date if date.is_a? Numeric
    if date.is_a? Time
      date = DateTime.new date.year, date.month, date.day,
                          date.hour, date.min, date.sec
    end
    value = date - @worksheet.workbook.date_base
    if date > LEAP_ERROR
      value += 1
    end
    value
  end
  def encode_rk value
    #  Bit  Mask        Contents
    #    0  0x00000001  0 = Value not changed 1 = Value is multiplied by 100
    #    1  0x00000002  0 = Floating-point value 1 = Signed integer value
    # 31-2  0xFFFFFFFC  Encoded value
    cent = 0
    int = 2
    higher = value * 100
    if higher == higher.to_i
      value = higher.to_i
      cent = 1
    end
    if value.is_a?(Integer)
      shifted = [value].pack 'l'
      ## I can't find a format for packing a little endian signed integer
      shifted.reverse! if @bigendian
      value, = shifted.unpack 'V'
      value <<= 2
    else
      # FIXME: precision of small numbers
      int = 0
      value, = [value].pack(EIGHT_BYTE_DOUBLE).unpack('x4V')
      value &= 0xfffffffc
    end
    value | cent | int
  end
  def name
    unicode_string @worksheet.name
  end
  def row_blocks
    # All cells in an Excel document are divided into blocks of 32 consecutive
    # rows, called Row Blocks. The first Row Block starts with the first used
    # row in that sheet. Inside each Row Block there will occur ROW records
    # describing the properties of the rows, and cell records with all the cell
    # contents in this Row Block.
    blocks = []
    @worksheet.reject do |row| row.empty? end.each_with_index do |row, idx|
      blocks << [] if idx % 32 == 0
      blocks.last << row
    end
    blocks
  end
  def size
    @io.size
  end
  def strings
    @worksheet.inject [] do |memo, row|
      strings = row.select do |cell| cell.is_a? String end
      memo.concat strings
    end
  end
  ##
  # Write a blank cell
  def write_blank row, idx
    write_cell :blank, row, idx
  end
  def write_bof
    data = [
      @biff_version, # BIFF version (always 0x0600 for BIFF8)
      0x0010,        # Type of the following data:
                     # 0x0005 = Workbook globals
                     # 0x0006 = Visual Basic module
                     # 0x0010 = Worksheet
                     # 0x0020 = Chart
                     # 0x0040 = Macro sheet
                     # 0x0100 = Workspace file
      @build_id,     # Build identifier
      @build_year,   # Build year
      0x000,         # File history flags
      0x006,         # Lowest Excel version that can read
                     # all records in this file
    ]
    write_op @bof, data.pack("v4V2")
  end
  ##
  # Write a cell with a Boolean or Error value
  def write_boolerr row, idx
    value = row[idx]
    type = 0
    numval = 0
    if value.is_a? Error
      type = 1
      numval = value.code
    elsif value
      numval = 1
    end
    data = [
      numval, # Boolean or error value (type depends on the following byte)
      type    # 0 = Boolean value; 1 = Error code
    ]
    write_cell :boolerr, row, idx, *data
  end
  def write_calccount
    count = 100 # Maximum number of iterations allowed in circular references
    write_op 0x000c, [count].pack('v')
  end
  def write_cell type, row, idx, *args
    xf_idx = @workbook.xf_index @worksheet.workbook, row.format(idx)
    data = [
      row.idx, # Index to row
      idx,     # Index to column
      xf_idx,  # Index to XF record (➜ 6.115)
    ].concat args
    write_op opcode(type), data.pack(binfmt(type))
  end
  def write_cellblocks row
    # BLANK ➜ 6.7
    # BOOLERR ➜ 6.10
    # INTEGER ➜ 6.56 (BIFF2 only)
    # LABEL ➜ 6.59 (BIFF2-BIFF7)
    # LABELSST ➜ 6.61 (BIFF8 only)
    # MULBLANK ➜ 6.64 (BIFF5-BIFF8)
    # MULRK ➜ 6.65 (BIFF5-BIFF8)
    # NUMBER ➜ 6.68
    # RK ➜ 6.82 (BIFF3-BIFF8)
    # RSTRING ➜ 6.84 (BIFF5/BIFF7)
    multiples, first_idx = nil
    row.each_with_index do |cell, idx|
      if multiples && (!multiples.last.is_a?(cell.class) \
                       || (cell.is_a?(Numeric) && cell.abs < 0.1))
        write_multiples row, first_idx, multiples
        multiples, first_idx = nil
      end
      nxt = idx + 1
      case cell
      when NilClass
        if multiples
          multiples.push cell
        elsif nxt < row.size && row[nxt].nil?
          multiples = [cell]
          first_idx = idx
        else
          write_blank row, idx
        end
      when TrueClass, FalseClass, Error
        write_boolerr row, idx
      when String
        write_labelsst row, idx
      when Numeric
        ## RK encodes Floats with 30 significant bits, which is a bit more than
        #  10^9. Not sure what is a good rule of thumb here, but it seems that
        #  Decimal Numbers with more than 4 significant digits are not represented
        #  with sufficient precision by RK
        if cell.is_a?(Float) && cell.to_s.length > 5
          write_number row, idx
        elsif multiples
          multiples.push cell
        elsif nxt < row.size && row[nxt].is_a?(Numeric)
          multiples = [cell]
          first_idx = idx
        else
          write_rk row, idx
        end
      when Formula
        write_formula row, idx
      when Date
        write_rk row, idx
      end
    end
    write_multiples row, first_idx, multiples if multiples
  end
  def write_changes reader, endpos, sst_status
    reader.seek @worksheet.offset
    blocks = row_blocks
    lastpos = reader.pos
    offsets = {}
    @worksheet.offsets.each do |key, pair|
      if @worksheet.changes.include?(key) \
        || (sst_status == :complete_update && key.is_a?(Integer))
        offsets.store pair, key
      end
    end
    offsets.invert.sort_by do |key, (pos, len)|
      pos
    end.each do |key, (pos, len)|
      @io.write reader.read(pos - lastpos)
      if key.is_a?(Integer)
        block = blocks.find do |rows| rows.any? do |row| row.idx == key end end
        write_rowblock block
      else
        send "write_#{key}"
      end
      lastpos = pos + len
      reader.seek lastpos
    end
    @io.write reader.read(endpos - lastpos)
  end
  def write_defaultrowheight
    data = [
      0x00, # Option flags:
            # Bit  Mask  Contents
            #   0  0x01  1 = Row height and default font height do not match
            #   1  0x02  1 = Row is hidden
            #   2  0x04  1 = Additional space above the row
            #   3  0x08  1 = Additional space below the row
      0xf2, #   Default height for unused rows, in twips = 1/20 of a point
    ]
    write_op 0x0225, data.pack('v2')
  end
  def write_dimensions
    # Offset  Size  Contents
    #      0     4  Index to first used row
    #      4     4  Index to last used row, increased by 1
    #      8     2  Index to first used column
    #     10     2  Index to last used column, increased by 1
    #     12     2  Not used
    write_op 0x0200, @worksheet.dimensions.pack(binfmt(:dimensions))
  end
  def write_eof
    write_op 0x000a
  end
  ##
  # Write a cell with a Formula. May write an additional String record depending
  # on the stored result of the Formula.
  def write_formula row, idx
    cell = row[idx]
    data1 = [
      row.idx,      # Index to row
      idx,          # Index to column
      0,            # Index to XF record (➜ 6.115)
    ].pack 'v3'
    data2 = nil
    case value = cell.value
    when Numeric    # IEEE 754 floating-point value (64-bit double precision)
      data2 = [value].pack EIGHT_BYTE_DOUBLE
    when String
      data2 = [
        0x00,       # (identifier for a string value)
        0xffff,     #
      ].pack 'Cx5v'
    when true, false
      value = value ? 1 : 0
      data2 = [
        0x01,     # (identifier for a Boolean value)
        value,    # 0 = FALSE, 1 = TRUE
        0xffff,   #
      ].pack 'CxCx3v'
    when Error
      data2 = [
        0x02,       # (identifier for an error value)
        value.code, # Error code
        0xffff,     #
      ].pack 'CxCx3v'
    when nil
      data2 = [
        0x03,       # (identifier for an empty cell)
        0xffff,     #
      ].pack 'Cx5v'
    else
      data2 = [
        0x02,       # (identifier for an error value)
        0x2a,       # Error code: #N/A! Argument or function not available
        0xffff,     #
      ].pack 'CxCx3v'
    end
    opts = 0x03
    opts |= 0x08 if cell.shared
    data3 = [
      opts        # Option flags:
                  # Bit  Mask    Contents
                  #   0  0x0001  1 = Recalculate always
                  #   1  0x0002  1 = Calculate on open
                  #   3  0x0008  1 = Part of a shared formula
    ].pack 'vx4'
    write_op opcode(:formula), data1, data2, data3, cell.data
    if cell.value.is_a?(String)
      write_op opcode(:string), unicode_string(cell.value, 2)
    end
  end
  ##
  # Write a new Worksheet.
  def write_from_scratch
    # ●  BOF Type = worksheet (➜ 6.8)
    write_bof
    # ○  UNCALCED ➜ 6.104
    # ○  INDEX ➜ 5.7 (Row Blocks), ➜ 6.55
    # ○  Calculation Settings Block ➜ 5.3
    write_calccount
    write_refmode
    write_iteration
    write_saverecalc
    # ○  PRINTHEADERS ➜ 6.76
    # ○  PRINTGRIDLINES ➜ 6.75
    # ○  GRIDSET ➜ 6.48
    # ○  GUTS ➜ 6.49
    # ○  DEFAULTROWHEIGHT ➜ 6.28
    write_defaultrowheight
    # ○  WSBOOL ➜ 6.113
    write_wsbool
    # ○  Page Settings Block ➜ 5.4
    # ○  Worksheet Protection Block ➜ 5.18
    # ○  DEFCOLWIDTH ➜ 6.29
    # ○○ COLINFO ➜ 6.18
    # ○  SORT ➜ 6.95
    # ●  DIMENSIONS ➜ 6.31
    write_dimensions
    # ○○ Row Blocks ➜ 5.7
    write_rows
    # ●  Worksheet View Settings Block ➜ 5.5
    # ○  STANDARDWIDTH ➜ 6.97
    # ○○ MERGEDCELLS ➜ 6.63
    # ○  LABELRANGES ➜ 6.60
    # ○  PHONETIC ➜ 6.73
    # ○  Conditional Formatting Table ➜ 5.12
    # ○  Hyperlink Table ➜ 5.13
    # ○  Data Validity Table ➜ 5.14
    # ○  SHEETLAYOUT ➜ 6.91 (BIFF8X only)
    # ○  SHEETPROTECTION Additional protection, ➜ 6.92 (BIFF8X only)
    # ○  RANGEPROTECTION Additional protection, ➜ 6.79 (BIFF8X only)
    # ●  EOF ➜ 6.36
    write_eof
  end
  def write_iteration
    its = 0 # 0 = Iterations off; 1 = Iterations on
    write_op 0x0011, [its].pack('v')
  end
  ##
  # Write a cell with a String value. The String must have been stored in the
  # Shared String Table.
  def write_labelsst row, idx
    write_cell :labelsst, row, idx, @workbook.sst_index(self, row[idx])
  end
  ##
  # Write multiple consecutive blank cells.
  def write_mulblank row, idx, multiples
    data = [
      row.idx, # Index to row
      idx, # Index to first column (fc)
    ]
    # List of nc=lc-fc+1 16-bit indexes to XF records (➜ 6.115)
    multiples.each_with_index do |blank, cell_idx|
      xf_idx = @workbook.xf_index @worksheet.workbook, row.format(idx + cell_idx)
      data.push xf_idx
    end
    # Index to last column (lc)
    data.push idx + multiples.size
    write_op opcode(:mulblank), data.pack('v*')
  end
  ##
  # Write multiple consecutive cells with RK values (see #write_rk)
  def write_mulrk row, idx, multiples
    fmt = 'v2'
    data = [
      row.idx, # Index to row
      idx, # Index to first column (fc)
    ]
    # List of nc=lc-fc+1 16-bit indexes to XF records (➜ 6.115)
    multiples.each do |cell|
      # TODO: XF indices
      data.push 0, encode_rk(cell)
      fmt << 'vV'
    end
    # Index to last column (lc)
    data.push idx + multiples.size
    write_op opcode(:mulrk), data.pack(fmt << 'v')
  end
  def write_multiples row, idx, multiples
    case multiples.last
    when NilClass
      write_mulblank row, idx, multiples
    when Numeric
      write_mulrk row, idx, multiples
    end
  end
  ##
  # Write a cell with a 64-bit double precision Float value
  def write_number row, idx
    # Offset Size Contents
    # 0 2 Index to row
    # 2 2 Index to column
    # 4 2 Index to XF record (➜ 6.115)
    # 6 8 IEEE 754 floating-point value (64-bit double precision)
    write_cell :number, row, idx, row[idx]
  end
  def write_op op, *args
    data = args.join
    @io.write [op,data.size].pack("v2")
    @io.write data
  end
  def write_refmode
    # • The “RC” mode uses numeric indexes for rows and columns, for example
    #   “R(1)C(-1)”, or “R1C1:R2C2”.
    # • The “A1” mode uses characters for columns and numbers for rows, for
    #   example “B1”, or “$A$1:$B$2”.
    mode = 1 # 0 = RC mode; 1 = A1 mode
    write_op 0x000f, [mode].pack('v')
  end
  ##
  # Write a cell with a Numeric or Date value.
  def write_rk row, idx
    value = row[idx]
    case value
    when Date, DateTime
      value = encode_date(value)
    end
    write_cell :rk, row, idx, encode_rk(value)
  end
  def write_row row
    # Offset  Size  Contents
    #      0     2  Index of this row
    #      2     2  Index to column of the first cell which
    #               is described by a cell record
    #      4     2  Index to column of the last cell which is
    #               described by a cell record, increased by 1
    #      6     2  Bit   Mask    Contents
    #               14-0  0x7fff  Height of the row, in twips = 1/20 of a point
    #                 15  0x8000  0 = Row has custom height;
    #                             1 = Row has default height
    #      8     2  Not used
    #     10     1  0 = No defaults written;
    #               1 = Default row attribute field and XF index occur below (fl)
    #     11     2  Relative offset to calculate stream position of the first
    #               cell record for this row (➜ 5.7.1)
    #   [13]     3  (written only if fl = 1) Default row attributes (➜ 3.12)
    #   [16]     2  (written only if fl = 1) Index to XF record (➜ 6.115)
    has_defaults = row.default_format ? 1 : 0
    data = [
      row.idx,
      row.first_used,
      row.first_unused,
      row.height * TWIPS,
      0, # Not used
      has_defaults,
      0, # OOffice does not set this - ignore until someone complains
    ]
    # OpenOffice apparently can't read Rows with a length other than 16 Bytes
    fmt = binfmt(:row) + 'x3'
=begin
    if format = row.default_format
      fmt = fmt + 'xv'
      data.concat [
        #0, # Row attributes should only matter in BIFF2
        workbook.xf_index(@worksheet.workbook, format),
      ]
    end
=end
    write_op opcode(:row), data.pack(fmt)
  end
  def write_rowblock block
    # ●● ROW Properties of the used rows
    # ○○ Cell Block(s) Cell records for all used cells
    # ○  DBCELL Stream offsets to the cell records of each row
    block.each do |row|
      write_row row
    end
    block.each do |row|
      write_cellblocks row
    end
  end
  def write_rows
    row_blocks.each do |block|
      write_rowblock block
    end
  end
  def write_saverecalc
    # 0 = Do not recalculate; 1 = Recalculate before saving the document
    write_op 0x005f, [1].pack('v')
  end
  def write_wsbool
    bits = [
         #   Bit  Mask    Contents
      1, #     0  0x0001  0 = Do not show automatic page breaks
         #                1 = Show automatic page breaks
      0, #     4  0x0010  0 = Standard sheet
         #                1 = Dialogue sheet (BIFF5-BIFF8)
      0, #     5  0x0020  0 = No automatic styles in outlines
         #                1 = Apply automatic styles to outlines
      1, #     6  0x0040  0 = Outline buttons above outline group
         #                1 = Outline buttons below outline group
      1, #     7  0x0080  0 = Outline buttons left of outline group
         #                1 = Outline buttons right of outline group
      0, #     8  0x0100  0 = Scale printout in percent (➜ 6.89)
         #                1 = Fit printout to number of pages (➜ 6.89)
      0, #     9  0x0200  0 = Save external linked values
         #                    (BIFF3-BIFF4 only, ➜ 5.10)
         #                1 = Do not save external linked values
         #                    (BIFF3-BIFF4 only, ➜ 5.10)
      1, #    10  0x0400  0 = Do not show row outline symbols
         #                1 = Show row outline symbols
      0, #    11  0x0800  0 = Do not show column outline symbols
         #                1 = Show column outline symbols
      0, # 13-12  0x3000  These flags specify the arrangement of windows.
         #                They are stored in BIFF4 only.
         #                00 = Arrange windows tiled
         #                01 = Arrange windows horizontal
      0, #                10 = Arrange windows vertical
         #                11 = Arrange windows cascaded
         # The following flags are valid for BIFF4-BIFF8 only:
      0, #    14  0x4000  0 = Standard expression evaluation
         #                1 = Alternative expression evaluation
      0, #    15  0x8000  0 = Standard formula entries
         #                1 = Alternative formula entries
    ]
    weights = [4,5,6,7,8,9,10,11,12,13,14,15]
    value = bits.inject do |a, b| a | (b << weights.shift) end
    write_op 0x0081, [value].pack('v')
  end
end
    end
  end
end
