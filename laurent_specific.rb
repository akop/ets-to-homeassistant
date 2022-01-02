lambda do |knx|
  # 1: fix special blinds...
  o_delete=[]
  o_new={}
  knx[:ob].each do |k,o|
    if o[:type].eql?(:sun_protection) and knx[:ga][o[:ga].first][:name].end_with?(':Pulse')
      # split this object into 2: delete old object
      o_delete.push(k)
      # create one obj per GA
      o[:ga].each do |gid|
        direction=case knx[:ga][gid][:name]
        when /Montee/;'Montee'
        when /Descente/;'Descente'
        else raise "error: #{knx[:ga][gid][:name]}"
        end
        # fix datapoint for ga
        knx[:ga][gid][:datapoint].replace('1.001')
        # create new object
        o_new["#{k}_#{direction}"]={
          name:   "#{o[:name]} #{direction}",
          type:   :custom, # simple switch
          ga:     [gid],
          floor:  o[:floor],
          room:   o[:room],
          custom: {ha_type: 'switch'} # custom values
        }
      end
    end
  end
  # delete redundant objects
  o_delete.each{|i|knx[:ob].delete(i)}
  # add split objects
  knx[:ob].merge!(o_new)
  # 2: set unique name when needed
  knx[:ob].each do |k,o|
    # set name as room + function
    o[:custom][:ha_init]={'name'=>"#{o[:name]} #{o[:room]}"}
    # set my specific parameters
    if o[:type].eql?(:sun_protection)
      o[:custom][:ha_init].merge!({
        'travelling_time_down'=>59,
        'travelling_time_up'=>59
      })
    end
  end
  # 3- manage group addresses without object
  error=false
  knx[:ga].values.select{|ga|ga[:objs].empty?}.each do |ga|
    error=true
    STDERR.puts("group not in object: #{ga[:address]}")
  end
  raise "Error found" if error
end