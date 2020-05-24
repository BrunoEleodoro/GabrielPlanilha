stamp := $(shell date +%F-%r)

build: 
	@echo "\033[0;32mStarting now... $(stamp)\033[0m" 
	@echo "\033[0;32mGenerating new spreadsheet...\033[0m" 
	node criar_planilha.js
	@echo "\033[0;32mGenerating month...\033[0m" 
	node rename_columns.js
	@echo "\033[0;32mRename columns...\033[0m" 
	node gerar_month.js
	@echo "\033[0;32mGenerating week_day...\033[0m" 
	node gerar_week_day.js
	@echo "\033[0;32mFiltering type...\033[0m" 
	node filtrar_type.js
	@echo "\033[0;32mFiltering client...\033[0m" 
	node filtrar_clientes.js
	@echo "\033[0;32mFiltering severities...\033[0m" 
	node filtrar_severidades.js
	@echo "\033[0;32mLabels based on severities...\033[0m" 
	node labels_based_sevs.js
	@echo "\033[0;32mFiltering labels...\033[0m" 
	node filtrar_labels.js
	@echo "\033[0;32mCalculate Hours for each ticket...\033[0m" 
	node calculate_hours.js
	@echo "\033[0;32mCalculate Amount of tickets for each person...\033[0m" 
	node count_and_amount_of_hours.js
	@echo "\033[0;32mFix the wrong dates for created at...\033[0m" 
	node corrigir_created_at.js
	@echo "\033[0;32mFix the general format to numbers parseFloat...\033[0m" 
	node corrigir_numeros.js
	@echo "\033[0;32mGenerating Tribe...\033[0m" 
	node gerar_tribe.js
	@echo "\033[0;32mGenerating Horario Pico...\033[0m" 
	node gerar_horario_pico.js
	@echo "\033[0;32mGenerating shift...\033[0m" 
	node filtrar_shift.js
	@echo "\033[0;32mSLA Ticket Vencido...\033[0m" 
	node sla_ticket_vencido.js
	@echo "\033[0;32mTempo Atendimento...\033[0m" 
	node tempo_atendimento.js
	@echo "\033[0;32mAnalise Prazo SLA...\033[0m" 
	node analise_prazo.js

	# @echo "\033[0;32mInverter datas...\033[0m" 
	# node inverter_datas.js
	# @echo "\033[0;32mLabels for each employee...\033[0m" 
	# node separar_tickets.js

consolidado: 
	@echo "\033[0;32mStarting now... $(stamp)\033[0m" 
	@echo "\033[0;32mGenerating month...\033[0m" 
	node gerar_month.js
	@echo "\033[0;32mGenerating week_day...\033[0m" 
	node gerar_week_day.js