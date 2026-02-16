run_testing: testing
	killall soffice.bin || echo "No libreoffice instance found"
	lowriter "$$TESTING_ODT" --norestore &

testing: src/vimbrewriter.vbs
	./compile.sh "src/vimbrewriter.vbs" "$$TESTING_XBA"

extension: clean src/vimbrewriter.vbs
	if [ -z "$$VIMBREWRITER_VERSION" ]; then \
		echo "VIMBREWRITER_VERSION must be set"; \
	else \
		mkdir -p build; mkdir -p dist; \
		cp -r extension/template build/template; \
		./compile.sh "src/vimbrewriter.vbs" "build/template/vimbrewriter/vimbrewriter.xba"; \
		cd "build/template"; \
		sed -i "s/%VIMBREWRITER_VERSION%/$$VIMBREWRITER_VERSION/g" description.xml; \
		7z a "../../dist/vimbrewriter-$$VIMBREWRITER_VERSION.oxt" .; \
	fi

.PHONY: clean
clean:
	rm -rf build