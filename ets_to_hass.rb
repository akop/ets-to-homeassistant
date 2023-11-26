#!/usr/bin/env ruby
# frozen_string_literal: true

# Laurent Martin
# translate configuration from ETS into KNXWeb and Home Assistant
require 'zip'
require 'xmlsimple'
require 'yaml'
require 'json'
require 'logger'
require 'getoptlong'

class ConfigurationImporter
  # extension of ETS project file
  ETS_EXT = '.knxproj'
  # converters of group address integer address into representation
  GADDR_CONV = {
    Free: ->(a) { a.to_s },
    TwoLevel: ->(a) { [(a >> 11) & 31, a & 2047].join('/') },
    ThreeLevel: ->(a) { [(a >> 11) & 31, (a >> 8) & 7, a & 255].join('/') }
  }.freeze
  # KNX functions described in knx_master.xml in project file.
  # map index parsed from "FT-x" to recognizable identifier
  ETS_FUNCTIONS = %i[custom switchable_light dimmable_light sun_protection heating_radiator heating_floor
                     dimmable_light sun_protection heating_switching_variable heating_continuous_variable].freeze
  # https://www.home-assistant.io/integrations/knx/#value-types
  ETS_GA_DATAPOINT_2_HA_SENSOR_ADDRESS_TYPE = {
    '5' => '1byte_unsigned',
    '5.001' => 'percent',
    '5.003' => 'angle',
    '5.004' => 'percentU8',
    '5.005' => 'decimal_factor',
    '5.006' => 'tariff',
    '5.010' => 'pulse',
    '6' => '1byte_signed',
    '6.001' => 'percentV8',
    '6.010' => 'counter_pulses',
    '7' => '2byte_unsigned',
    '7.001' => 'pulse_2byte',
    '7.002' => 'time_period_msec',
    '7.003' => 'time_period_10msec',
    '7.004' => 'time_period_100msec',
    '7.005' => 'time_period_sec',
    '7.006' => 'time_period_min',
    '7.007' => 'time_period_hrs',
    '7.011' => 'length_mm',
    '7.012' => 'current',
    '7.013' => 'brightness',
    '7.600' => 'color_temperature',
    '8' => '2byte_signed',
    '8.001' => 'pulse_2byte_signed',
    '8.002' => 'delta_time_ms',
    '8.003' => 'delta_time_10ms',
    '8.004' => 'delta_time_100ms',
    '8.005' => 'delta_time_sec',
    '8.006' => 'delta_time_min',
    '8.007' => 'delta_time_hrs',
    '8.010' => 'percentV16',
    '8.011' => 'rotation_angle',
    '8.012' => 'length_m',
    '9' => '2byte_float',
    '9.001' => 'temperature',
    '9.002' => 'temperature_difference_2byte',
    '9.003' => 'temperature_a',
    '9.004' => 'illuminance',
    '9.005' => 'wind_speed_ms',
    '9.006' => 'pressure_2byte',
    '9.007' => 'humidity',
    '9.008' => 'ppm',
    '9.009' => 'air_flow',
    '9.010' => 'time_1',
    '9.011' => 'time_2',
    '9.020' => 'voltage',
    '9.021' => 'curr',
    '9.022' => 'power_density',
    '9.023' => 'kelvin_per_percent',
    '9.024' => 'power_2byte',
    '9.025' => 'volume_flow',
    '9.026' => 'rain_amount',
    '9.027' => 'temperature_f',
    '9.028' => 'wind_speed_kmh',
    '9.029' => 'absolute_humidity',
    '9.030' => 'concentration_ugm3',
    '9.?' => 'enthalpy',
    '12' => '4byte_unsigned',
    '12.001' => 'pulse_4_ucount',
    '12.100' => 'long_time_period_sec',
    '12.101' => 'long_time_period_min',
    '12.102' => 'long_time_period_hrs',
    '12.1200' => 'volume_liquid_litre',
    '12.1201' => 'volume_m3',
    '13' => '4byte_signed',
    '13.001' => 'pulse_4byte',
    '13.002' => 'flow_rate_m3h',
    '13.010' => 'active_energy',
    '13.011' => 'apparant_energy',
    '13.012' => 'reactive_energy',
    '13.013' => 'active_energy_kwh',
    '13.014' => 'apparant_energy_kvah',
    '13.015' => 'reactive_energy_kvarh',
    '13.016' => 'active_energy_mwh',
    '13.100' => 'long_delta_timesec',
    '14' => '4byte_float',
    '14.000' => 'acceleration',
    '14.001' => 'acceleration_angular',
    '14.002' => 'activation_energy',
    '14.003' => 'activity',
    '14.004' => 'mol',
    '14.005' => 'amplitude',
    '14.006' => 'angle_rad',
    '14.007' => 'angle_deg',
    '14.008' => 'angular_momentum',
    '14.009' => 'angular_velocity',
    '14.010' => 'area',
    '14.011' => 'capacitance',
    '14.012' => 'charge_density_surface',
    '14.013' => 'charge_density_volume',
    '14.014' => 'compressibility',
    '14.015' => 'conductance',
    '14.016' => 'electrical_conductivity',
    '14.017' => 'density',
    '14.018' => 'electric_charge',
    '14.019' => 'electric_current',
    '14.020' => 'electric_current_density',
    '14.021' => 'electric_dipole_moment',
    '14.022' => 'electric_displacement',
    '14.023' => 'electric_field_strength',
    '14.024' => 'electric_flux',
    '14.025' => 'electric_flux_density',
    '14.026' => 'electric_polarization',
    '14.027' => 'electric_potential',
    '14.028' => 'electric_potential_difference',
    '14.029' => 'electromagnetic_moment',
    '14.030' => 'electromotive_force',
    '14.031' => 'energy',
    '14.032' => 'force',
    '14.033' => 'frequency',
    '14.034' => 'angular_frequency',
    '14.035' => 'heatcapacity',
    '14.036' => 'heatflowrate',
    '14.037' => 'heat_quantity',
    '14.038' => 'impedance',
    '14.039' => 'length',
    '14.040' => 'light_quantity',
    '14.041' => 'luminance',
    '14.042' => 'luminous_flux',
    '14.043' => 'luminous_intensity',
    '14.044' => 'magnetic_field_strength',
    '14.045' => 'magnetic_flux',
    '14.046' => 'magnetic_flux_density',
    '14.047' => 'magnetic_moment',
    '14.048' => 'magnetic_polarization',
    '14.049' => 'magnetization',
    '14.050' => 'magnetomotive_force',
    '14.051' => 'mass',
    '14.052' => 'mass_flux',
    '14.053' => 'momentum',
    '14.054' => 'phaseanglerad',
    '14.055' => 'phaseangledeg',
    '14.056' => 'power',
    '14.057' => 'powerfactor',
    '14.058' => 'pressure',
    '14.059' => 'reactance',
    '14.060' => 'resistance',
    '14.061' => 'resistivity',
    '14.062' => 'self_inductance',
    '14.063' => 'solid_angle',
    '14.064' => 'sound_intensity',
    '14.065' => 'speed',
    '14.066' => 'stress',
    '14.067' => 'surface_tension',
    '14.068' => 'common_temperature',
    '14.069' => 'absolute_temperature',
    '14.070' => 'temperature_difference',
    '14.071' => 'thermal_capacity',
    '14.072' => 'thermal_conductivity',
    '14.073' => 'thermoelectric_power',
    '14.074' => 'time_seconds',
    '14.075' => 'torque',
    '14.076' => 'volume',
    '14.077' => 'volume_flux',
    '14.078' => 'weight',
    '14.079' => 'work',
    '14.080' => 'apparent_power',
    '16.000' => 'string',
    '16.001' => 'latin_1',
    '17.001' => 'scene_number'
  }.freeze
  private_constant :ETS_EXT, :ETS_FUNCTIONS, :ETS_GA_DATAPOINT_2_HA_SENSOR_ADDRESS_TYPE

  attr_reader :data

  def initialize(file, options = {})
    # set to true if the resulting yaml starts at the knx key
    @opts = options
    # parsed data: ob: objects, ga: group addresses
    @data = { ob: {}, ga: {} }
    # log to stderr, so that redirecting stdout captures only generated data
    @logger = Logger.new($stderr)
    @logger.level = @opts.key?(:trace) ? @opts[:trace] : Logger::INFO
    @logger.debug("options: #{@opts}")
    project = read_file(file)
    proj_info = self.class.dig_xml(project[:info], %w[Project ProjectInformation])
    group_addr_style = @opts.key?(:addr) ? @opts[:addr] : proj_info['GroupAddressStyle']
    @logger.info("Using project #{proj_info['Name']}, address style: #{group_addr_style}")
    # set group address formatter according to project settings
    @addrparser = GADDR_CONV[group_addr_style.to_sym]
    raise "Error: no such style #{group_addr_style} in #{GADDR_CONV.keys}" if @addrparser.nil?

    installation = self.class.dig_xml(project[:data], %w[Project Installations Installation])
    # process group ranges: fill @data[:ga]
    process_group_ranges(self.class.dig_xml(installation, %w[GroupAddresses GroupRanges]))
    # process group ranges: fill @data[:ob] (for 2 versions of ETS which have different tags?)
    process_space(self.class.dig_xml(installation, ['Locations']), 'Space') if installation.key?('Locations')
    process_space(self.class.dig_xml(installation, ['Buildings']), 'BuildingPart') if installation.key?('Buildings')
    @logger.warn('No building information found.') if @data[:ob].keys.empty?
  end

  # helper function to dig through keys, knowing that we used ForceArray
  def self.dig_xml(entry_point, path)
    raise "ERROR: wrong entry point: #{entry_point.class}, expect Hash" unless entry_point.is_a?(Hash)

    path.each do |n|
      raise "ERROR: cannot find level #{n} in xml, have #{entry_point.keys.join(',')}" unless entry_point.key?(n)

      entry_point = entry_point[n]
      # because we use ForceArray
      entry_point = entry_point.first
      raise "ERROR: expect array with one element in #{n}" if entry_point.nil?
    end
    entry_point
  end

  # Read both project.xml and 0.xml
  # @return Hash {info: xmldata, data: xmldata}
  def read_file(file)
    raise "ETS file must end with #{ETS_EXT}" unless file.end_with?(ETS_EXT)

    project = {}
    # read ETS5 file and get project file
    Zip::File.open(file) do |zip_file|
      zip_file.each do |entry|
        case entry.name
        when %r{P-[^/]+/project\.xml$}
          project[:info] = XmlSimple.xml_in(entry.get_input_stream.read, { 'ForceArray' => true })
        when %r{P-[^/]+/0\.xml$}
          project[:data] = XmlSimple.xml_in(entry.get_input_stream.read, { 'ForceArray' => true })
        end
      end
    end
    raise "Did not find project information or data (#{project.keys})" unless project.keys.sort.eql?(%i[data info])

    project
  end

  # process group range recursively and find addresses
  def process_group_ranges(gr)
    gr['GroupRange'].each { |sgr| process_group_ranges(sgr) } if gr.key?('GroupRange')
    gr['GroupAddress'].each { |ga| process_ga(ga) } if gr.key?('GroupAddress')
  end

  # process a group address
  def process_ga(ga)
    # build object for each group address
    group = {
      name: ga['Name'].freeze, # ETS: name field
      description: ga['Description'].freeze, # ETS: description field
      address: @addrparser.call(ga['Address'].to_i).freeze, # group address as string. e.g. "x/y/z" depending on project style
      datapoint: nil, # datapoint type as string "x.00y"
      objs: [], # objects ids, it may be in multiple objects
      custom: {} # modified by lambda
    }
    if ga['DatapointType'].nil?
      @logger.warn("no datapoint type for #{group[:address]} : #{group[:name]}, group address is skipped")
      return
    end
    # parse datapoint for easier use
    if (m = ga['DatapointType'].match(/^DPST-([0-9]+)-([0-9]+)$/))
      # datapoint type as string x.00y
      group[:datapoint] = format('%d.%03d', m[1].to_i, m[2].to_i) # no freeze
    else
      @logger.warn("Cannot parse data point type: #{ga['DatapointType']}, group is skipped, expect: DPST-x-x")
      return
    end
    # Index is the internal Id in xml file
    @data[:ga][ga['Id'].freeze] = group.freeze
    @logger.debug("group: #{group}")
  end

  # process locations recursively, and find functions
  # @param space the current space
  # @param info current location information: floor, room
  def process_space(space, space_type, info = nil)
    @logger.debug(">sname>#{space['Type']}: #{space['Name']}")
    @logger.debug(">space>#{space}")
    info = info.nil? ? {} : info.dup
    # process building sub spaces
    if space.key?(space_type)
      # get floor when we have it
      info[:floor] = space['Name'] if space['Type'].eql?('Floor')
      # get buildingpart when we have it
      info[:buildingpart] = space['Name'] if space['Type'].eql?('BuildingPart')
      space[space_type].each { |s| process_space(s, space_type, info) }
    end
    # Functions are objects
    return unless space.key?('Function')

    # we assume the object is directly in the room
    info[:room] = space['Name']
    # loop on group addresses
    space['Function'].each do |f|
      @logger.debug("function #{f}")
      # ignore functions without group address
      next unless f.key?('GroupAddressRef')

      m = f['Type'].match(/^FT-([0-9])$/)
      raise "ERROR: Unknown function type: #{f['Type']}" if m.nil?

      type = ETS_FUNCTIONS[m[1].to_i]

      # the object
      o = {
        name: f['Name'].freeze,
        type:,
        ga: f['GroupAddressRef'].map { |g| g['RefId'].freeze },
        custom: {} # custom values
      }.merge(info)
      # store reference to this object in the GAs
      o[:ga].each { |g| @data[:ga][g][:objs].push(f['Id']) if @data[:ga].key?(g) }
      @logger.debug("function: #{o}")
      @data[:ob][f['Id']] = o.freeze
    end
  end

  def generate_homeass
    haknx = {}
    # warn of group addresses that will not be used (you can fix in custom lambda)
    @data[:ga].values.select { |ga| ga[:objs].empty? }.each do |ga|
      @logger.warn("group not in object: #{ga[:address]}: Create custom object in lambda if needed , or use ETS to create functions")
    end
    @data[:ob].each_value do |o|
      new_obj = o[:custom].key?(:ha_init) ? o[:custom][:ha_init] : {}
      unless new_obj.key?('name')
        new_obj['name'] = if true && @opts[:full_name]
                            "#{o[:buildingpart]} #{o[:floor]} #{o[:room]} #{o[:name]}"
                          else
                            o[:name]
                          end
      end
      # compute object type
      ha_obj_type =
        if o[:custom].key?(:ha_type)
          o[:custom][:ha_type]
        else
          # map FT-x type to home assistant type
          case o[:type]
          when :switchable_light, :dimmable_light then 'light'
          when :sun_protection then 'cover'
          when :custom
            if not o[:name].include?('|')
              @logger.warn("#{o[:name]}/#{o[:room]}: #{o[:type]}: custom dynamic function type not implemented with '<my ETS function name> |<ha_obj_type>' (example: |sensor or |binary_sensor)")
              next
            end
            # get ha_type from custom name: "my awesome |switch" -> switch
            pos_from = o[:name].rindex("|") + 1
            custom_dynamic_ha_obj_type = o[:name][pos_from..-1]
            # remove custom dynamic ha_type from new object name
            new_name = new_obj['name'].gsub('|' + custom_dynamic_ha_obj_type, '').strip
            new_obj['name'].replace(new_name)
            custom_dynamic_ha_obj_type
          when :heating_continuous_variable, :heating_floor, :heating_radiator, :heating_switching_variable
            @logger.warn("function type not implemented for #{o[:name]}/#{o[:room]}: #{o[:type]}")
            next
          else @logger.error("function type not supported for #{o[:name]}/#{o[:room]}, please report: #{o[:type]}")
            next
          end
        end
      # process all group addresses in function
      o[:ga].each do |garef|
        ga = @data[:ga][garef]
        next if ga.nil?

        # find property name based on datapoint
        ha_address_type =
          if ga[:custom].key?(:ha_address_type)
            ga[:custom][:ha_address_type]
          #
          # ha_address_type mapping for custom ga based on ga name: 'my name |<ha_address_type>'
          # for example: "my switchable socket |state_address"
          #
          elsif (o[:type].eql?(:custom) && (ga[:name].include? '|'))
            pos_from = ga[:name].rindex("|") + 1
            custom_dynamic_ha_address_type = ga[:name][pos_from..-1]
            custom_dynamic_ha_address_type
          #
          # Add sensor type specific gas based on its datapoint
          #
          elsif (ha_obj_type == 'sensor')
            new_obj['type'] = ETS_GA_DATAPOINT_2_HA_SENSOR_ADDRESS_TYPE[ga[:datapoint]]
            'state_address'
          else
            case ga[:datapoint]
            when '1.001' then 'address' # switch on/off or state
            when '1.008' then 'move_long_address' # up/down
            when '1.010' then 'stop_address' # stop
            when '1.011' then 'state_address' # switch state
            when '3.007'
              @logger.debug("#{ga[:address]}(#{ha_obj_type}:#{ga[:datapoint]}:#{ga[:name]}): ignoring datapoint")

              next # dimming control: used by buttons
            when '5.001' # percentage 0-100
              # custom code tells what is state
              case ha_obj_type
              when 'light'
                if ! ga[:name].include? '|'
                  @logger.warn("#{ga[:address]}(#{ha_obj_type}:#{ga[:datapoint]}:#{ga[:name]}): missing custom dynamic light ha address type in ga name 'my address name |<ha_address_type>' (example: |brightness_address or |brightness_state_address)")
                  next
                end
                pos_from = ga[:name].rindex("|") + 1
                custom_dynamic_light_ha_address_type = ga[:name][pos_from..-1]
                custom_dynamic_light_ha_address_type
              when 'cover'
                if ! ga[:name].include? '|'
                  @logger.warn("#{ga[:address]}(#{ha_obj_type}:#{ga[:datapoint]}:#{ga[:name]}): missing custom dynamic cover ha address type in ga name 'my address name |<ha_address_type>' (example: |position_address or |position_state_address)")
                  next
                end
                pos_from = ga[:name].rindex("|") + 1
                custom_dynamic_cover_ha_address_type = ga[:name][pos_from..-1]
                custom_dynamic_cover_ha_address_type
              else @logger.warn("#{ga[:address]}(#{ha_obj_type}:#{ga[:datapoint]}:#{ga[:name]}): no mapping for datapoint #{ga[:datapoint]}")
                   next
              end
            else
              @logger.warn("#{ga[:address]}(#{ha_obj_type}:#{ga[:datapoint]}:#{ga[:name]}): no mapping for datapoint #{ga[:datapoint]}")

              next
            end
          end
                if ha_address_type.nil?
          @logger.warn("#{ga[:address]}(#{ha_obj_type}:#{ga[:datapoint]}:#{ga[:name]}): unexpected nil property name")
          next
        end
        if new_obj.key?(ha_address_type)
          @logger.error("#{ga[:address]}(#{ha_obj_type}:#{ga[:datapoint]}:#{ga[:name]}): ignoring for #{ha_address_type} already set with #{new_obj[ha_address_type]}")
          next
        end
        # Add addtional move_short_address for cover which points to stop_address ga
        if (ha_obj_type == 'cover' && ha_address_type == 'stop_address')
          new_obj['move_short_address'] = ga[:address]
        end
        new_obj[ha_address_type] = ga[:address]
      end
      haknx[ha_obj_type] = [] unless haknx.key?(ha_obj_type)
      # check name is not duplicated, as name is used to identify the object
      if haknx[ha_obj_type].any? { |v| v['name'].casecmp?(new_obj['name']) }
        @logger.warn("object name is duplicated: #{new_obj['name']}")
      end
      haknx[ha_obj_type].push(new_obj)
    end
    return { 'knx' => haknx }.to_yaml if @opts[:ha_knx]

    haknx.to_yaml
  end

  # https://sourceforge.net/p/linknx/wiki/Object_Definition_section/
  def generate_linknx
    @data[:ga].values.sort { |a, b| a[:address] <=> b[:address] }.map do |ga|
      linknx_disp_name = ga[:custom][:linknx_disp_name] || ga[:name]
      %(        <object type="#{ga[:datapoint]}" id="id_#{ga[:address].gsub('/',
                                                                            '_')}" gad="#{ga[:address]}" init="request">#{linknx_disp_name}</object>)
    end.join("\n")
  end
end

# prefix of generation methods
GENPREFIX = 'generate_'
# get list of generation methods
genformats = (ConfigurationImporter.instance_methods - ConfigurationImporter.superclass.instance_methods)
             .select { |m| m.to_s.start_with?(GENPREFIX) }
             .map { |m| m[GENPREFIX.length..-1] }

opts = GetoptLong.new(
  ['--help', '-h', GetoptLong::NO_ARGUMENT],
  ['--format', '-f', GetoptLong::REQUIRED_ARGUMENT],
  ['--ha-knx', '-k', GetoptLong::NO_ARGUMENT],
  ['--full-name', '-n', GetoptLong::NO_ARGUMENT],
  ['--lambda', '-l', GetoptLong::REQUIRED_ARGUMENT],
  ['--trace', '-t', GetoptLong::REQUIRED_ARGUMENT]
)

options = {}

custom_lambda = File.join(File.dirname(__FILE__), 'default_custom.rb')
output_format = 'homeass'
opts.each do |opt, arg|
  case opt
  when '--help'
    puts <<-EOF
            Usage: #{$PROGRAM_NAME} [--format format] [--lambda lambda] [--addr addr] [--trace trace] [--ha-knx] [--full-name] <etsprojectfile>.knxproj

            -h, --help:
              show help

            --format [format]:
              one of #{genformats.join('|')}

            --lambda [lambda]:
              file with lambda

            --addr [addr]:
              one of #{ConfigurationImporter::GADDR_CONV.keys.map(&:to_s).join(', ')}

            --trace [trace]:
              one of debug, info, warn, error

            --ha-knx:
              include level knx in ouput file

            --full-name:
              add room name in object name
    EOF
    Process.exit(1)
  when '--lambda'
    custom_lambda = arg
  when '--format'
    output_format = arg
    raise "Error: no such output format: #{output_format}" unless genformats.include?(output_format)
  when '--ha-knx'
    options[:ha_knx] = true
  when '--full-name'
    options[:full_name] = true
  when '--trace'
    options[:trace] = arg
  else
    raise "Unknown option #{opt}"
  end
end

if ARGV.length != 1
  puts 'Missing project file argument (try --help)'
  Process.exit(1)
end

infile = ARGV.shift

# read and parse ETS file
knxconf = ConfigurationImporter.new(infile, options)
# apply special code if provided
eval(File.read(custom_lambda), binding, custom_lambda).call(knxconf) unless custom_lambda.nil?
# generate result
$stdout.write(knxconf.send("#{GENPREFIX}#{output_format}".to_sym))
