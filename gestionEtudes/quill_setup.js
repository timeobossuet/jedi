let quillContexte, quillProblematique, quillDescription, quillPhase, quillClientDesc;

function initQuillEditors() {
    const quillOptions = {
        theme: 'snow',
        modules: {
            toolbar: [
                ['bold', 'italic', 'underline', 'strike'],
                [{ 'list': 'ordered'}, { 'list': 'bullet' }],
                ['clean']
            ]
        }
    };
    
    // Contexte
    quillContexte = new Quill('#etude-contexte', quillOptions);
    quillContexte.on('text-change', function(delta, oldDelta, source) {
        if (source === 'user') {
            mettreAJourInfoEtude('contexte', quillContexte.root.innerHTML);
        }
    });

    // Problématique
    quillProblematique = new Quill('#etude-problematique', quillOptions);
    quillProblematique.on('text-change', function(delta, oldDelta, source) {
        if (source === 'user') {
            mettreAJourInfoEtude('problematique', quillProblematique.root.innerHTML);
        }
    });

    // Notes Rapides / Description
    quillDescription = new Quill('#etude-description', quillOptions);
    quillDescription.on('text-change', function(delta, oldDelta, source) {
        if (source === 'user') {
            document.getElementById('btn-save-desc').classList.remove('opacity-0');
        }
    });

    // Phases
    quillPhase = new Quill('#phase-editor-description', quillOptions);
}

function initClientQuill() {
    const quillOptions = {
        theme: 'snow',
        modules: {
            toolbar: [
                ['bold', 'italic', 'underline', 'strike'],
                [{ 'list': 'ordered'}, { 'list': 'bullet' }],
                ['clean']
            ]
        }
    };
    if (document.getElementById('client-description-editor')) {
        quillClientDesc = new Quill('#client-description-editor', quillOptions);
        if (currentClientInfo && currentClientInfo.description) {
            quillClientDesc.root.innerHTML = currentClientInfo.description;
        }
        quillClientDesc.on('text-change', function(delta, oldDelta, source) {
            if (source === 'user') {
                mettreAJourChampInfoClient('description', quillClientDesc.root.innerHTML);
            }
        });
    }
}
