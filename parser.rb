# This class will perform the actual parsing

class Parser
  require 'unzipper'
  require 'libxml'
  attr_accessor :file_name, :string, :rtf, :html, :txt
  
  
  def initialize(file_name)
    @file_name = file_name
    @basename = File.basename(file_name, File.extname(file_name))
    @string = ""
    @rtf = ""
    UnZipper.unzip(@file_name)
  end
  
  # to_raw_string just returns the text with basic string formatting. By default, it includes spaces (line breaks and indents).
  def to_raw_string(preserve_spacing = true)
    doc = LibXML::XML::Document.file(@basename << '/word/document.xml')
    path = doc.find('/w:document/w:body')
    node = path.first
    if @string.empty?
      raw_string_digger(node, preserve_spacing)
    end
    puts @string    
  end

  # this method uses recursion to dig deep into an xml tree
  # it checks for text, line breaks, indentions, and tabs
  def raw_string_digger (node, preserve_spacing)
    # raw text
    if node.text? && !node.to_s.nil?
        @string << node.to_s
    end
    # line breaks
    if node.name == "p" && !node.empty?
      @string << "\n" if preserve_spacing && !@string.empty?
      @string << " " if !preserve_spacing && !@string.empty?
    end
    if preserve_spacing
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
      raw_string_digger(child, preserve_spacing)
    end
  end

  def to_rtf(preserve_spacing = true)
    doc = LibXML::XML::Document.file(@basename << '/word/document.xml')
    path = doc.find('/w:document/w:body')
    node = path.first
    if true ##@rtf.empty?
      # Set the header to specify ansi, black Arial, and other paramaters...
      @rtf = %q/{\rtf1\ansi\ansicpg1252\cocoartf1138\cocoasubrtf320
{\fonttbl\f0\fswiss\fcharset0 ArialMT;}
{\colortbl;\red255\green255\blue255;}
\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\deftab720
\pard\pardeftab720\ri720/
      @memo = {:bold => '', :itallic => '', :underline => ''}
      rtf_digger(node, preserve_spacing)
    end
    @rtf << "}"
    puts @rtf
  end
  
  # does everything that the raw_string_digger does and more
  # also checks for bold, underline, and itaics
  # of course converting to the correct rtf encoding
  def rtf_digger (node, preserve_spacing)
    # raw text
    if node.text? && !node.to_s.nil?
        @rtf << node.to_s
    end
    # line breaks
    if node.name == "p" && !node.empty?
      @rtf << %q/\ / if preserve_spacing && !@string.empty?
      @rtf << " " if !preserve_spacing && !@string.empty?
    end
    # start of <w:r> (not sure what this stands for. "row" maybe?)
    # this means we should reset memo (all previous bold, itallics, or underline tags should end)
    if node.name == "r"
      puts @memo.inspect
      @rtf << @memo[:bold] << @memo[:itallic] << @memo[:underline]
      @memo = {:bold => '', :itallic => '', :underline => ''}
    end
    #bold
    if node.name == 'b'
      @rtf << %q/ \b /
      @memo[:bold] = %q/\b0/ # so we remember to close the bold tag later on
    end
    #italics
    if node.name == 'i'
      @rtf << %q/ \i /
      @memo[:bold] = %q/\i0/ # so we remember to close the itallics tag later on
    end
    #underline
    if node.name == 'u'
      @rtf << %q/ \ul /
      @memo[:bold] = %q/\ul0/ # so we remember to close the underline tag later on
    end
    if preserve_spacing
      # indentions (assuming an indent value of 720 corresponds to one tab)
      if node.name == "ind"
        tab_num = (node.attributes.first.value.to_i)/720
        tab_num.times{@rtf << "\t"}
      end
      # tabs
      if node.name == "tab"
        @rtf << "\t"
      end
    end
    # isn't recursion fun? ...
    node.children.each do |child|      
      rtf_digger(child, preserve_spacing)
    end
  end
    
end

# this is used for testing. The test.docx uses italics, bold, underline, and several fonts
parser = Parser.new('test.docx')
parser.to_rtf
