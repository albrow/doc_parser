# This class will perform the actual parsing

class Parser
  require 'unzipper'
  require 'libxml'
  attr_accessor :file_name, :string, :rtf, :html, :txt
  
  
  def initialize(file_name)
    @file_name = file_name
    @basename = File.basename(file_name, File.extname(file_name))
    @string = ""
    UnZipper.unzip(@file_name)
  end
  
  # to_raw_string just returns the text with basic string formatting. By default, it includes spaces (line breaks and indents).
  def to_raw_string(spaces = true)
    doc = LibXML::XML::Document.file(@basename << '/word/document.xml')
    path = doc.find('/w:document/w:body')
    node = path.first
    if @string.empty?
      raw_string_digger(node, spaces)
    end
    puts @string    
  end

  # this method uses recursion to dig deep into an xml tree
  # it checks for text, line breaks, indentions, and tabs
  def raw_string_digger (node, spaces)
    # raw text
    if node.text? && !node.to_s.nil?
      @string << node.to_s
    end
    # line breaks
    if node.name == "p" && !node.empty?
      @string << "\n" if spaces && !@string.empty?
      @string << " " if !spaces && !@string.empty?
    end
    # the following should happen only if spaces is enabled...
    if spaces
      # indentions (assuming an indent value of 720 corresponds to one tab)
      if node.name == "ind"
        tab_num = (node.attributes.first.value.to_i)/720
        tab_num.times{@string << "\t"}
      end
      # tabs
      if node.name == "tab"
        @string << "\t"
      end
    end
    # isn't recursion fun? ...
    node.children.each do |child|      
      raw_string_digger(child, spaces)
    end
  end

    
end

# this is used for testing. The test.docx uses italics, bold, underline, and several fonts
parser = Parser.new('test.docx')
parser.to_raw_string
