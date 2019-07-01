stamp := $(shell date +%F-%r)

build: 
	@echo "\033[0;32mStarting now... $(stamp)\033[0m" 
	@echo "\033[0;32mGenerating new spreadsheet...\033[0m" 
	node criar_planilha.js
	@echo "\033[0;32mGenerating week_day...\033[0m" 
	node gerar_week_day.js
	@echo "\033[0;32mGenerating month...\033[0m" 
	node gerar_month.js
	@echo "\033[0;32mGenerating shift...\033[0m" 
	node filtrar_shift.js
	@echo "\033[0;32mFiltering type...\033[0m" 
	node filtrar_type.js
	@echo "\033[0;32mFiltering client...\033[0m" 
	node filtrar_clientes.js
	@echo "\033[0;32mFiltering severities...\033[0m" 
	node filtrar_severidades.js
	@echo "\033[0;32mFiltering labels...\033[0m" 
	node filtrar_labels.js