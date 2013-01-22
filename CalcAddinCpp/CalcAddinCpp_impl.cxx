/*	Serhat Sevki Dincer, jfcgaussATgmail, 12apr2006
	Random Number Generation Add-In for OpenOffice.org Calc

	This program can be used, modified and/or (re)distributed under
	the terms of LGPL. See www.gnu.org for info about LGPL.

	This program has NO warranty.
	OpenOffice.org SDK has been used while writing this program.
*/


#include <stdio.h>
#include <wchar.h>
#include <sys/time.h>
#include <cstdlib>
#include <cmath>

#include <cppuhelper/implbase4.hxx> //4-parameter template will be used
#include <cppuhelper/factory.hxx>
#include <cppuhelper/implementationentry.hxx>

#include <com/sun/star/sheet/XAddIn.hpp>
#include <com/sun/star/lang/XServiceName.hpp>
#include <com/sun/star/lang/XServiceInfo.hpp>
#include <org/openoffice/sheet/addin/XCalcAddinCpp.hpp>

#include <sal/main.h>

#include <cppuhelper/bootstrap.hxx>

#include <com/sun/star/beans/XPropertySet.hpp>
#include <com/sun/star/frame/XComponentLoader.hpp>
#include <com/sun/star/lang/XMultiComponentFactory.hpp>
#include "com/sun/star/lang/XMultiServiceFactory.hpp"
#include <com/sun/star/registry/XSimpleRegistry.hpp>

// scalc
#include <com/sun/star/sheet/XSpreadsheet.hpp>
#include <com/sun/star/table/XCell.hpp>
#include <com/sun/star/sheet/XSpreadsheetDocument.hpp>
#include <com/sun/star/sheet/XSpreadsheets.hpp>
#include <com/sun/star/sheet/XSpreadsheetDocument.hpp>

#include <com/sun/star/style/XStyleFamiliesSupplier.hpp>

#include <com/sun/star/table/CellRangeAddress.hpp>
#include <com/sun/star/table/XCell.hpp>
#include <com/sun/star/table/XCellRange.hpp>
#include <com/sun/star/table/XTableChart.hpp>
#include <com/sun/star/table/XTableCharts.hpp>
#include <com/sun/star/table/XTableChartsSupplier.hpp>

#include "com/sun/star/container/XNameAccess.hpp"
#include "com/sun/star/container/XIndexAccess.hpp"

using namespace ::rtl;
using namespace ::com::sun::star;
using namespace ::com::sun::star::uno;

using namespace com::sun::star::lang;
using namespace com::sun::star::beans;
using namespace com::sun::star::frame;
using namespace com::sun::star::registry;

// scalc
using namespace ::com::sun::star::sheet;
using namespace ::com::sun::star::style;
using namespace ::com::sun::star::table;
using namespace ::com::sun::star::container;

using ::rtl::OUString;
using ::rtl::OUStringToOString;

static const   int X = 10;
static const   int Y = 50;

namespace _CalcAddinCpp_impl_
{
  /// SCalc code ported to C++

void insertIntoCellF(int CellX, int CellY, float theValue, Reference < XSpreadsheet > TT1)
{
    try {
        Reference < XCell > xCell = TT1->getCellByPosition(CellX, CellY);
        xCell->setValue(theValue);
    } catch (IndexOutOfBoundsException ex) {
        printf("Could not get Cell\n");
    }
}

void insertIntoCellS(int CellX, int CellY, OUString theValue, Reference < XSpreadsheet > TT1)
{
    try {
        Reference < XCell > xCell = TT1->getCellByPosition(CellX, CellY);
        xCell->setFormula(theValue);
    } catch (IndexOutOfBoundsException ex) {
        printf("Could not get Cell\n");
    }
}

void chgbColor ( int x1, int y1, int x2, int y2,OUString tmplate, Reference < XSpreadsheet > TT )
{
    try {
        Reference < XCellRange > xCR = TT->getCellRangeByPosition(x1,y1,x2,y2);
        Reference < XPropertySet > xCPS ( xCR, UNO_QUERY );
        xCPS->setPropertyValue("CellStyle", Any(tmplate));
    } catch (IndexOutOfBoundsException ex) {
        printf("Could not get CellRange\n");
    } catch (Exception e) {
        printf("Can't change colors chgbColor\n");
    }
}

void createHeader ( Reference < XSpreadsheet > xSheet)
{
    printf("Creating the Header\n");

    insertIntoCellS(1,0,"JAN",xSheet);
    insertIntoCellS(2,0,"FEB",xSheet);
    insertIntoCellS(3,0,"MAR",xSheet);
    insertIntoCellS(4,0,"APR",xSheet);
    insertIntoCellS(5,0,"MAI",xSheet);
    insertIntoCellS(6,0,"JUN",xSheet);
    insertIntoCellS(7,0,"JUL",xSheet);
    insertIntoCellS(8,0,"AUG",xSheet);
    insertIntoCellS(9,0,"SEP",xSheet);
    insertIntoCellS(10,0,"OCT",xSheet);
    insertIntoCellS(11,0,"NOV",xSheet);
    insertIntoCellS(12,0,"DEC",xSheet);
    insertIntoCellS(13,0,"SUM",xSheet);
}

void fillLines ( Reference < XSpreadsheet > xSheet)
{
    printf("Fill the lines\n");

    Reference < XCell > xCell;
    for (int ii=0; ii < 120; ++ii)
        {
            for (int jj=0; jj < X; ++jj)
                {
                    for (int kk=0; kk < Y; ++kk)
                        {
                            xCell = xSheet->getCellByPosition(jj+1, kk+1);
                            xCell->setValue(jj+kk+ii+2);
                        }
                }
        }

    //***************************************************************************
}

void applyStyle ( Reference < XSpreadsheet > xSheet)
{
    //oooooooooooooooooooooooooooStep 5oooooooooooooooooooooooooooooooooooooooooo
    // apply the created cell style.
    // For this purpose get the PropertySet of the Cell and change the
    // property CellStyle to the appropriate value.
    //***************************************************************************

    // change backcolor
    chgbColor( 1 , 0, 13, 0, "My Style", xSheet );
    chgbColor( 0 , 1, 0, 3, "My Style", xSheet );
    chgbColor( 1 , 1, 13, 3, "My Style2", xSheet );

    //***************************************************************************
}

void genSpread (Reference< XSpreadsheetDocument > myDoc)
{
    //oooooooooooooooooooooooooooStep 3oooooooooooooooooooooooooooooooooooooooooo
    // create cell styles.
    // For this purpose get the StyleFamiliesSupplier and the the familiy
    // CellStyle. Create an instance of com.sun.star.style.CellStyle and
    // add it to the family. Now change some properties
    //***************************************************************************

    try {
        Reference < XStyleFamiliesSupplier > xSFS (myDoc, UNO_QUERY);
                
        Reference < XNameAccess > xSF = xSFS->getStyleFamilies();
        Reference < XMultiServiceFactory > oDocMSF (myDoc, UNO_QUERY);
        Reference < XNameContainer > oStyleFamilyNameContainer (xSF->getByName("CellStyles"), UNO_QUERY);
        Reference < XInterface > oInt1 = oDocMSF->createInstance("com.sun.star.style.CellStyle");
        oStyleFamilyNameContainer->insertByName("My Style", Any(oInt1));
        Reference < XPropertySet > oCPS1 ( oInt1, UNO_QUERY );
        oCPS1->setPropertyValue("IsCellBackgroundTransparent", Any(sal_False));
        oCPS1->setPropertyValue("CellBackColor",Any(6710932));
        oCPS1->setPropertyValue("CharColor",Any(16777215));
        Reference < XInterface > oInt2 = oDocMSF->createInstance("com.sun.star.style.CellStyle");
        oStyleFamilyNameContainer->insertByName("My Style2", Any(oInt2));
        Reference < XPropertySet > oCPS2 ( oInt2, UNO_QUERY );
        oCPS2->setPropertyValue("IsCellBackgroundTransparent", Any(sal_False));
        oCPS2->setPropertyValue("CellBackColor",Any(13421823));

        printf("setting style on spreadsheet\n");
    } catch (Exception e) {
        printf("*** Error setting style on spreadsheet\n");
    }

    //***************************************************************************

    //oooooooooooooooooooooooooooStep 4oooooooooooooooooooooooooooooooooooooooooo
    // get the sheet an insert some data.
    // Get the sheets from the document and then the first from this container.
    // Now some data can be inserted. For this purpose get a Cell via
    // getCellByPosition and insert into this cell via setValue() (for floats)
    // or setFormula() for formulas and Strings
    //***************************************************************************

    try {
        printf("Getting spreadsheet\n");
        Reference < XSpreadsheets > xSheets = myDoc->getSheets() ;
        Reference < XIndexAccess > oIndexSheets ( xSheets, UNO_QUERY);
        Any any = oIndexSheets->getByIndex(0);
        Reference < XSpreadsheet > xSheet ( any, UNO_QUERY);

        createHeader (xSheet);
        fillLines (xSheet);
        applyStyle (xSheet);

    } catch (Exception e) {
        printf("*** Error Couldn't get Sheet\n");
    }

}

void getDoc (void)
{
  // get the remote office component context
  Reference< XComponentContext > xContext( ::cppu::bootstrap() );

  /* Gets the service manager instance to be used (or null). This method has
	 been added for convenience, because the service manager is a often used
	 object.
  */
  Reference< XMultiComponentFactory > xMCFC(xContext->getServiceManager() );

  /* Creates an instance of a component which supports the services specified
	 by the factory. Important: using the office component context.
  */
  Reference< XInterface > xIface = xMCFC->createInstanceWithContext(OUString("com.sun.star.frame.Desktop"),xContext );
  Reference< XComponentLoader > xComponentLoader(xIface, UNO_QUERY );

  // gets the server component context as property of the office component factory
  Reference< XPropertySet > xPropSet( xIface, UNO_QUERY );

  OUString sAbsoluteDocUrl = OUString ("private:factory/scalc"); // uncomment for local scalc
  Reference< XComponent > xComponent = xComponentLoader->loadComponentFromURL(
																			  sAbsoluteDocUrl, OUString( "_default" ), 0, // _blank for new doc, _default for exitsing
																			  Sequence < ::com::sun::star::beans::PropertyValue >() );

  Reference< XSpreadsheetDocument > myDoc (xComponent, UNO_QUERY); 
  if (myDoc.is())
	{
	  printf("spreadsheet sAbsoluteDocUrl=%s\n", OUStringToOString(sAbsoluteDocUrl, RTL_TEXTENCODING_ASCII_US).getStr());
	  genSpread(myDoc);
	}
  else
	{
	  printf("sAbsoluteDocUrl=%s\n", OUStringToOString(sAbsoluteDocUrl, RTL_TEXTENCODING_ASCII_US).getStr());
	}
}
  /// SCalc code ported to C++

	class CalcAddinCpp_impl : public ::cppu::WeakImplHelper4< ::org::openoffice::sheet::addin::XCalcAddinCpp,
		sheet::XAddIn, lang::XServiceName, lang::XServiceInfo > //4-parameter template
	{
	//Locale
		lang::Locale locale;

	public:
	//XCalcAddinCpp
	  double	SAL_CALL expo( double m ) throw (RuntimeException); //expo(mean)
	  int		SAL_CALL getMyFirstValue( void ) throw (RuntimeException); //1st value
	  OUString  SAL_CALL methodOne( OUString const & str ) throw (RuntimeException);
	  OUString  SAL_CALL methodTwo( OUString const & str ) throw (RuntimeException);
	  sal_Int32 SAL_CALL methodThree(const Sequence< Sequence< sal_Int32 > > &aValList ) throw (RuntimeException);
	  Sequence< Sequence< sal_Int32 > > 
	            SAL_CALL methodFour( const Sequence< Sequence< sal_Int32 > > &aValList ) throw (RuntimeException);

	//XAddIn
		OUString SAL_CALL getProgrammaticFuntionName( const OUString& aDisplayName ) throw (RuntimeException);
		OUString SAL_CALL getDisplayFunctionName( const OUString& aProgrammaticName ) throw (RuntimeException);
		OUString SAL_CALL getFunctionDescription( const OUString& aProgrammaticName ) throw (RuntimeException);
		OUString SAL_CALL getDisplayArgumentName( const OUString& aProgrammaticName, ::sal_Int32 nArgument ) throw (RuntimeException);
		OUString SAL_CALL getArgumentDescription( const OUString& aProgrammaticName, ::sal_Int32 nArgument ) throw (RuntimeException);
		OUString SAL_CALL getProgrammaticCategoryName( const OUString& aProgrammaticName ) throw (RuntimeException);
		OUString SAL_CALL getDisplayCategoryName( const OUString& aProgrammaticName ) throw (RuntimeException);

	//XServiceName
		OUString SAL_CALL getServiceName(  ) throw (RuntimeException);

	//XServiceInfo
		OUString SAL_CALL getImplementationName(  ) throw (RuntimeException);
		::sal_Bool SAL_CALL supportsService( const OUString& ServiceName ) throw (RuntimeException);
		Sequence< OUString > SAL_CALL getSupportedServiceNames(  ) throw (RuntimeException);

	//XLocalizable
		void SAL_CALL setLocale( const lang::Locale& eLocale ) throw (RuntimeException);
		lang::Locale SAL_CALL getLocale(  ) throw (RuntimeException);

		short  SAL_CALL getFunctionID( const OUString& aDisplayName );

	};

//XCalcAddinCpp
	//expo(mean)
	double CalcAddinCpp_impl::expo( double m ) throw (RuntimeException)
	{
		return -m * log( (double)(1+(unsigned int)rand()) / (2+(unsigned int)RAND_MAX) );
	}
	int	   CalcAddinCpp_impl::getMyFirstValue(void ) throw (RuntimeException)
	{
	  return 1;
	}
  OUString CalcAddinCpp_impl::methodOne( OUString const & str )
    throw (RuntimeException)
  {
    return OUString(RTL_CONSTASCII_USTRINGPARAM("called methodOne() of MyService2 implementation: ") ) + str;
  }
 
  OUString CalcAddinCpp_impl::methodTwo( OUString const & str )
    throw (RuntimeException)
  {
	getDoc();
    return OUString( RTL_CONSTASCII_USTRINGPARAM("called methodTwo() of MyService2 implementation: ") ) + str;
  }
 
 
  sal_Int32 CalcAddinCpp_impl::methodThree(const Sequence< Sequence< sal_Int32 > > &aValList )
    throw (RuntimeException)
  {
  	sal_Int32		n1, n2;
  	sal_Int32		nE1 = aValList.getLength();
  	sal_Int32		nE2;
  	sal_Int32 temp=0;
  	for( n1 = 0 ; n1 < nE1 ; n1++ )
  	  {
  		const Sequence< sal_Int32 >	rList = aValList[ n1 ];
  		nE2 = rList.getLength();
  		const sal_Int32*	pList = rList.getConstArray();
  		for( n2 = 0 ; n2 < nE2 ; n2++ )
  		  {
  			temp += pList[ n2 ];
  		  }
  	  }
  	return temp;
  }
  //It's a matrix operation  should be called like : {=METHODFOUR(A1:B4)}
  Sequence< Sequence< sal_Int32 > > CalcAddinCpp_impl::methodFour(const Sequence< Sequence< sal_Int32 > > &aValList )throw (RuntimeException)
  {
  	sal_Int32		n1, n2;
  	sal_Int32		nE1 = aValList.getLength();
  	sal_Int32		nE2;
  	Sequence< Sequence< sal_Int32 > > temp = aValList;
  	for( n1 = 0 ; n1 < nE1 ; n1++ )
  	  {
  		Sequence< sal_Int32 >	rList = temp[ n1 ];
  		nE2 = rList.getLength();
  		for( n2 = 0 ; n2 < nE2 ; n2++ )
  		  {
  			rList[ n2 ] += 4;
  		  }
  		temp[n1]=rList;
  	  }
  	return temp;
  }


	#define _serviceName_ "org.openoffice.sheet.addin.CalcAddinCpp"

	static const sal_Char *_serviceName = _serviceName_;
    static const int	   _numFunc = 6;
    static const sal_Char*  _functionName[] = { "EXPO", "GETMYFIRSTVALUE", "methodOne", "methodTwo", "methodThree", "methodFour"};
    static const sal_Char* _functionDisplayName[] = {"expo", "getMyFirstValue", "methodOne", "methodTwo", "methodThree", "methodFour"};
    static const sal_Char* _functionDescription[] = {"random number", "my 1st value", "One", "Two", "Three", "Four"};
    static const sal_Char* _functionArgument[] = {"m", "none", "none", "none", "cells", "seq cells"};

//XAddIn
	OUString CalcAddinCpp_impl::getProgrammaticFuntionName( const OUString& aDisplayName ) throw (RuntimeException)
	{
	  short index = getFunctionID(aDisplayName);
	  if (index >= _numFunc) 	  return OUString(RTL_CONSTASCII_USTRINGPARAM("ERR"));

	  return OUString::createFromAscii(_functionName[index]);
	}

	OUString CalcAddinCpp_impl::getDisplayFunctionName( const OUString& aProgrammaticName ) throw (RuntimeException)
	{
	  short index = getFunctionID(aProgrammaticName);
	  if (index >= _numFunc) 	  return OUString(RTL_CONSTASCII_USTRINGPARAM("ERR"));

	  return OUString::createFromAscii(_functionDisplayName[index]);
	}

	OUString CalcAddinCpp_impl::getFunctionDescription( const OUString& aProgrammaticName ) throw (RuntimeException)
	{
	  short index = getFunctionID(aProgrammaticName);
	  if (index >= _numFunc) 	  return OUString(RTL_CONSTASCII_USTRINGPARAM("ERR"));

	  return OUString::createFromAscii(_functionDescription[index]);
	}

	OUString CalcAddinCpp_impl::getDisplayArgumentName( const OUString& aProgrammaticName, ::sal_Int32 nArgument ) throw (RuntimeException)
	{
	  short index = getFunctionID(aProgrammaticName);
	  if (index >= _numFunc) 	  return OUString(RTL_CONSTASCII_USTRINGPARAM("ERR"));

	  return OUString::createFromAscii(_functionArgument[index]);
	}

	OUString CalcAddinCpp_impl::getArgumentDescription( const OUString& aProgrammaticName, ::sal_Int32 nArgument ) throw (RuntimeException)
	{
	  short index = getFunctionID(aProgrammaticName);
	  if (index >= _numFunc) 	  return OUString(RTL_CONSTASCII_USTRINGPARAM("ERR"));

	  return OUString::createFromAscii(_functionArgument[index]);
	}

	OUString CalcAddinCpp_impl::getProgrammaticCategoryName( const OUString& aProgrammaticName ) throw (RuntimeException)
	{
		return OUString(RTL_CONSTASCII_USTRINGPARAM("Add-In"));
	}

	OUString CalcAddinCpp_impl::getDisplayCategoryName( const OUString& aProgrammaticName ) throw (RuntimeException)
	{
		return OUString(RTL_CONSTASCII_USTRINGPARAM("Add-In"));
	}

//XServiceName
	OUString CalcAddinCpp_impl::getServiceName(  ) throw (RuntimeException)
	{
		return OUString(_serviceName, sizeof(_serviceName_)-1, RTL_TEXTENCODING_ASCII_US);
	}

//XServiceInfo
	static OUString getImplementationName_CalcAddinCpp_impl() throw (RuntimeException)
	{
		return OUString(RTL_CONSTASCII_USTRINGPARAM("org.openoffice.sheet.addin.CalcAddinCpp_impl.CalcAddinCpp"));
	}

	OUString CalcAddinCpp_impl::getImplementationName() throw (RuntimeException)
	{
		return getImplementationName_CalcAddinCpp_impl();
	}

	::sal_Bool CalcAddinCpp_impl::supportsService( OUString const & serviceName ) throw (RuntimeException)
	{
		return serviceName.equalsAsciiL(_serviceName, sizeof(_serviceName_)-1);
	}

	static Sequence< OUString > getSupportedServiceNames_CalcAddinCpp_impl() throw (RuntimeException)
	{
		Sequence< OUString > name(1);
		name[0] = OUString(_serviceName, sizeof(_serviceName_)-1, RTL_TEXTENCODING_ASCII_US);
		return name;
	}

	Sequence< OUString > CalcAddinCpp_impl::getSupportedServiceNames() throw (RuntimeException)
	{
		return getSupportedServiceNames_CalcAddinCpp_impl();
	}

//XLocalizable
	void CalcAddinCpp_impl::setLocale( const lang::Locale& eLocale ) throw (RuntimeException)
	{
		locale = eLocale;
	}

	lang::Locale CalcAddinCpp_impl::getLocale(  ) throw (RuntimeException)
	{
		return locale;
	}

    short CalcAddinCpp_impl::getFunctionID (const OUString& stringProgmmaticName)
    {
	  for (int ii=0; ii < _numFunc; ++ii)
		{
		  const sal_Char* name = _functionDisplayName[ii];
		  if(stringProgmmaticName.equalsAscii(name))
			{
			  return (short) ii;
			}
		}
	  return -1;
    }

	static Reference< XInterface > SAL_CALL create_CalcAddinCpp_impl( Reference< XComponentContext > const & xContext ) SAL_THROW( () )
	{
		return static_cast< ::cppu::OWeakObject * > ( new CalcAddinCpp_impl );
	}

	static struct ::cppu::ImplementationEntry s_component_entries[] =
	{
		{ create_CalcAddinCpp_impl, getImplementationName_CalcAddinCpp_impl,
			getSupportedServiceNames_CalcAddinCpp_impl, ::cppu::createSingleComponentFactory, 0, 0 }
		,
		{ 0, 0, 0, 0, 0, 0 }
	};
}

extern "C"
{
	void SAL_CALL component_getImplementationEnvironment( sal_Char const ** ppEnvTypeName, uno_Environment ** ppEnv )
	{
		*ppEnvTypeName = CPPU_CURRENT_LANGUAGE_BINDING_NAME;
	}

	sal_Bool SAL_CALL component_writeInfo( lang::XMultiServiceFactory * xMgr, registry::XRegistryKey * xRegistry )
	{
		return ::cppu::component_writeInfoHelper( xMgr, xRegistry, ::_CalcAddinCpp_impl_::s_component_entries );
	}

	void * SAL_CALL component_getFactory( sal_Char const * implName,
		lang::XMultiServiceFactory * xMgr, registry::XRegistryKey * xRegistry )
	{
		return ::cppu::component_getFactoryHelper(implName, xMgr, xRegistry, ::_CalcAddinCpp_impl_::s_component_entries );
	}
}
