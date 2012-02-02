# This class will perform the actual parsing

class Parser
  require 'unzipper'
  require 'libxml'
  attr_accessor :file_name, :string, :rtf, :html, :txt
  
  
  def initialize(name)
    @file_name = name
    @string = ""
    UnZipper.unzip(@file_name)
  end
  
  # to_raw_string just gets the text with no formatting. By default, it includes line breaks.
  def to_raw_string(line_breaks = true)
    doc = LibXML::XML::Document.file('unzipped/word/document.xml')
    #puts doc
    path = doc.find('/w:document/w:body')
    node = path.first
    if @string.empty?
      raw_string_digger(node, line_breaks)
    end
    puts @string    
  end

  # this method uses recursion to dig deep into an xml tree
  # it checks for text and line breaks, but nothing else
  def raw_string_digger (node, line_breaks)
    # raw text
    if node.text? && !node.to_s.nil?
      @string << node.to_s
    end
    # line breaks
      if node.name == "p"
        @string << "\n" if line_breaks && !@string.empty?
        @string << " " if !line_breaks && !@string.empty?
      end
    node.children.each do |child|      
        raw_string_digger(child, line_breaks)
    end
  end

    
end

# this is used for testing. The test.docx uses italics, bold, underline, and several fonts
parser = Parser.new('test.docx')
parser.to_raw_string
