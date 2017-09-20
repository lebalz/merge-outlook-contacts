unless ARGV.size == 1
  raise "no valid path to outlook-exported file was provided.\n
        Call script with 'ruby #{__FILE__} /path/to/contacts.csv'"
end
raise "File #{ARGV[0]} does not exist!" unless File.exist?(ARGV[0])

FNAME = ARGV[0]

MAX_FIELDS = 61
FORMAT = 'bom|utf-8'.freeze

def i(name)
  name = name.lower.to_sym if name.is_a? String
  @header.index(name)
end

def conform_csv_contact(contact)
  contact.map! { |c| c.delete('"').strip.chomp }
  if contact.size == MAX_FIELDS
    contact
  elsif contact.size == MAX_FIELDS + 1
    contact[0..MAX_FIELDS - 2] + [contact[MAX_FIELDS - 1], contact[MAX_FIELDS]]
  else
    puts "contact had size of #{contact.size} and will not be imported.\n
          raw contact: #{contact.join(', ')}"
    nil
  end
end

def overlap?(entry1, entry2)
  !(entry1.nil? || entry1.empty? || entry1 == entry2)
end

def similar_field?(field_a, field_b)
  field_a.to_s.start_with?(field_b.to_s[0..4]) ||
    field_a.to_s.end_with?(field_b.to_s[-5..-1])
end

def first_similar_field(field)
  @header.find { |name| similar_field?(name, field) }
end

def next_similar_field(field)
  ind = i(field)
  @header[ind + 1] if similar_field?(field, @header[ind + 1])
end

def name_field?(contact)
  name_fields = %i[first_name middle_name last_name
                   nickname given_yomi surname_yomi]
  name_fields.all? { |field| contact[field].nil? || contact[field].empty? }
end

def handle_overlaps(overlays, merged_contact)
  overlays.each do |field, entries|
    next_field = first_similar_field(field)
    entries.each do |entry|
      while overlap?(merged_contact[next_field], entry)
        next_field = next_similar_field(next_field)
        break if next_field.nil?
      end
      if next_field.nil?
        puts "Field #{field} could not be merged and will be lost.
              #{merged_contact[:first_name]} #{merged_contact[:last_name]}:
              #{entry}"
      else
        merged_contact[next_field] = entry
      end
    end
  end
end

filtered = []
contacts = []
total_contacts = 0
File.open(FNAME, "r:#{FORMAT}") do |f|
  @header_raw = f.readline.split(',').map { |h| h.strip.chomp }
  @header = @header_raw.map do |column|
    column.delete('\'').gsub(%r{\s|-|\/}, '_').downcase.to_sym
  end
  f.each_line do |l|
    contact = conform_csv_contact(l.split(','))
    if contact
      contacts << @header.zip(contact).to_h.reject { |_, entry| entry.empty? }
    end
    total_contacts += 1
  end
end

grouped_contacts = contacts
                   .reject { |contact| name_field?(contact) }
                   .group_by { |contact| contact[:first_name].to_s + contact[:last_name].to_s }

grouped_contacts.each_value do |same_contacts|
  merged_contact = {}
  overlays = {}
  same_contacts.each do |contact|
    contact.each do |field, entry|
      if overlap?(merged_contact[field], entry)
        overlays[field] ||= []
        overlays[field] << entry unless overlays[field].include?(contact[field])
      else
        merged_contact[field] = entry
      end
    end
  end
  handle_overlaps(overlays, merged_contact)

  filtered << merged_contact
end
File.open(FNAME.split('.').first + '_merged.csv', 'w:utf-8', bom: true) do |f|
  f.puts @header_raw.join(',')
  filtered.each do |contact|
    f.puts @header.map { |field| contact[field] }.join(',')
  end
end
puts "merged #{total_contacts} contacts to #{filtered.size} contacts."
