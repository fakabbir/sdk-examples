
all: regcomp.flag

CalcAddinCpp.urd:
	$(OO_SDK_HOME)/bin/idlc -C -I$(OO_SDK_HOME)/idl CalcAddinCpp.idl

CalcAddinCpp.rdb: CalcAddinCpp.urd
	$(OO_SDK_URE_HOME)/bin/regmerge CalcAddinCpp.rdb /UCR CalcAddinCpp.urd

CalcAddinCpp_impl.o: CalcAddinCpp.rdb
	$(OO_SDK_HOME)/bin/cppumaker -BUCR -Torg.openoffice.sheet.addin.XCalcAddinCpp \
		-Tcom.sun.star.sheet.XAddIn \
		-Tcom.sun.star.lang.XServiceName \
		-Tcom.sun.star.lang.XServiceInfo \
		-Tcom.sun.star.lang.XTypeProvider \
		-Tcom.sun.star.uno.XWeak \
		-Tcom.sun.star.uno.XAggregation \
		-Tcom.sun.star.lang.XMultiServiceFactory \
		-Tcom.sun.star.uno.XComponentContext \
		-Tcom.sun.star.lang.XSingleComponentFactory \
		-Tcom.sun.star.lang.XSingleServiceFactory \
		-Tcom.sun.star.registry.XRegistryKey \
		$(OOo)/offapi.rdb \
		$(OO_SDK_URE_HOME)/share/misc/types.rdb \
		CalcAddinCpp.rdb
	g++ -O2 -fPIC -fomit-frame-pointer -c -o CalcAddinCpp_impl.o -DUNX \
		-DGCC -DLINUX -DCPPU_ENV=gcc3 -I. -I$(OO_SDK_HOME)/include -I $(OO_SDK_OUT)/LINUXexample.out/inc \
		CalcAddinCpp_impl.cxx

libCalcAddinCpp.so: CalcAddinCpp_impl.o
	ld -o libCalcAddinCpp.so -shared --retain-symbols-file symbols.txt \
	-L$(OO_SDK_URE_HOME)/lib -luno_cppuhelpergcc3 CalcAddinCpp_impl.o

regcomp.flag: libCalcAddinCpp.so
	$(OO_SDK_URE_HOME)/bin/regcomp -register -r CalcAddinCpp.rdb -c libCalcAddinCpp.so
	touch regcomp.flag

install : all
# /usr/bin/zip -u CalcAddinCpp.oxt  CalcAddinCpp.components META-INF/manifest.xml description.xml 
	/usr/bin/install -p libCalcAddinCpp.so CalcAddinCpp.rdb ${OFFICE_PROGRAM_PATH}/
	/usr/bin/zip -u CalcAddinCpp.oxt  CalcAddinCpp.components META-INF/manifest.xml CalcAddinCpp.rdb libCalcAddinCpp.so
	${OFFICE_PROGRAM_PATH}/unopkg add -f CalcAddinCpp.oxt

clean:
	rm -rf com org *.o *.urd *.so *.rdb *.flag *.oxt

