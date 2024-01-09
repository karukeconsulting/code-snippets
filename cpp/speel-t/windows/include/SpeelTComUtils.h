#include <msopc.h>
#include <comdef.h>

#define CHK_HR_THROW(stmt) do { HRESULT hr=(stmt); if (FAILED(hr)) { throw _com_error(hr); }} while(0)

_COM_SMARTPTR_TYPEDEF(IOpcFactory, IID_IOpcFactory);
_COM_SMARTPTR_TYPEDEF(IOpcPackage, IID_IOpcPackage);

_COM_SMARTPTR_TYPEDEF(IOpcPart, IID_IOpcPart);
_COM_SMARTPTR_TYPEDEF(IOpcPartUri, IID_IOpcPartUri);
_COM_SMARTPTR_TYPEDEF(IOpcPartSet, IID_IOpcPartSet);
_COM_SMARTPTR_TYPEDEF(IOpcPartEnumerator, IID_IOpcPartEnumerator);

_COM_SMARTPTR_TYPEDEF(IOpcRelationship, IID_IOpcRelationship);
_COM_SMARTPTR_TYPEDEF(IOpcRelationshipSet, IID_IOpcRelationshipSet);
_COM_SMARTPTR_TYPEDEF(IOpcRelationshipEnumerator, IID_IOpcRelationshipEnumerator);

_COM_SMARTPTR_TYPEDEF(IOpcUri, IID_IOpcUri);
_COM_SMARTPTR_TYPEDEF(IUri, IID_IUri);