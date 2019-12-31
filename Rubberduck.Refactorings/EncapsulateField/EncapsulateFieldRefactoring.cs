﻿using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.Utility;
using System.Collections.Generic;
using System;
using Antlr4.Runtime;
using Rubberduck.Refactorings.EncapsulateField.Extensions;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldRefactoringTestAccess
    {
        EncapsulateFieldModel TestUserInteractionOnly(Declaration target, Func<EncapsulateFieldModel, EncapsulateFieldModel> userInteraction);
    }

    public class EncapsulateFieldRefactoring : InteractiveRefactoringBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>, IEncapsulateFieldRefactoringTestAccess
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IIndenter _indenter;
        private QualifiedModuleName _targetQMN;
        private EncapsulateFieldElementFactory _encapsulationCandidateFactory;

        private enum NewContentTypes { TypeDeclarationBlock, DeclarationBlock, MethodBlock, PostContentMessage };
        private Dictionary<NewContentTypes, List<string>> _newContent { set; get; }

        private int? _codeSectionStartIndex;

        private static string DoubleSpace => $"{Environment.NewLine}{Environment.NewLine}";

        public EncapsulateFieldRefactoring(
            IDeclarationFinderProvider declarationFinderProvider,
            IIndenter indenter,
            IRefactoringPresenterFactory factory,
            IRewritingManager rewritingManager,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            IUiDispatcher uiDispatcher)
        :base(rewritingManager, selectionProvider, factory, uiDispatcher)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _indenter = indenter;
        }

        public EncapsulateFieldModel Model { set; get; }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            var selectedDeclaration = _selectedDeclarationProvider.SelectedDeclaration(targetSelection);
            if (selectedDeclaration == null
                || selectedDeclaration.DeclarationType != DeclarationType.Variable
                || selectedDeclaration.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member))
            {
                return null;
            }

            return selectedDeclaration;
        }

        public EncapsulateFieldModel TestUserInteractionOnly(Declaration target, Func<EncapsulateFieldModel, EncapsulateFieldModel> userInteraction)
        {
            var model = InitializeModel(target);
            return userInteraction(model);
        }

        protected override EncapsulateFieldModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            if (!target.DeclarationType.Equals(DeclarationType.Variable))
            {
                throw new InvalidDeclarationTypeException(target);
            }

            _targetQMN = target.QualifiedModuleName;

            var validator = new EncapsulateFieldValidator(_declarationFinderProvider) as IEncapsulateFieldValidator;
            _encapsulationCandidateFactory = new EncapsulateFieldElementFactory(_declarationFinderProvider, _targetQMN, validator);

            var candidates = _encapsulationCandidateFactory.CreateEncapsulationCandidates();
            var selected = candidates.Single(c => c.Declaration == target);
            selected.EncapsulateFlag = true;

            var forceUseOfObjectStateUDT = false;
            if (TryRetrieveExistingObjectStateUDT(target, candidates, out var objectStateUDT))
            {
                objectStateUDT.IsSelected = true;
                forceUseOfObjectStateUDT = true;
            }

            var defaultStateUDT = _encapsulationCandidateFactory.CreateStateUDTField();
            defaultStateUDT.IsSelected = objectStateUDT is null;

            Model = new EncapsulateFieldModel(
                                target,
                                candidates,
                                defaultStateUDT,
                                PreviewRewrite,
                                validator);

            if (forceUseOfObjectStateUDT)
            {
                Model.ConvertFieldsToUDTMembers = true;
                Model.StateUDTField = objectStateUDT;
            }

            _codeSectionStartIndex = _declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN).Where(m => m.IsMember())
                .OrderBy(c => c.Selection)
                            .FirstOrDefault()?.Context.Start.TokenIndex ?? null;

            return Model;
        }

        //Identify an existing objectStateUDT and make it unavailable for the user to select for encapsulation.
        //This prevents the user from inadvertently nesting a stateUDT within a new stateUDT
        private bool TryRetrieveExistingObjectStateUDT(Declaration target, IEnumerable<IEncapsulateFieldCandidate> candidates, out IObjectStateUDT objectStateUDT)
        {
            objectStateUDT = null;
            //Determination relies on matching the refactoring-generated name and a couple other UDT attributes
            //to determine if an objectStateUDT pre-exists the refactoring.

            //Question: would using an Annotations (like '@IsObjectStateUDT) be better?
            //The logic would then be: if Annotated => it's the one.  else => apply the matching criteria below
            
            //e.g., In cases where the user chooses an existing UDT for the initial encapsulation, the matching 
            //refactoring will not assign the name and the criteria below will fail => so applying an Annotation would
            //make it possible to find again
            var objectStateUDTIdentifier = $"{EncapsulateFieldResources.StateUserDefinedTypeIdentifierPrefix}{target.QualifiedModuleName.ComponentName}";

            var objectStateUDTMatches = candidates.Where(c => c is IUserDefinedTypeCandidate udt
                    && udt.Declaration.HasPrivateAccessibility()
                    && udt.Declaration.AsTypeDeclaration.IdentifierName.StartsWith(objectStateUDTIdentifier, StringComparison.InvariantCultureIgnoreCase))
                    .Select(pm => pm as IUserDefinedTypeCandidate);

            if (objectStateUDTMatches.Count() == 1)
            {
                objectStateUDT = new ObjectStateUDT(objectStateUDTMatches.First()) { IsSelected = true };
            }
            return objectStateUDT != null;
        }

        protected override void RefactorImpl(EncapsulateFieldModel model)
        {
            var refactorRewriteSession = new EncapsulateFieldRewriteSession(RewritingManager.CheckOutCodePaneSession()) as IEncapsulateFieldRewriteSession;
            refactorRewriteSession = RefactorRewrite(model, refactorRewriteSession);

            if (!refactorRewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(refactorRewriteSession.RewriteSession);
            }
        }

        private string PreviewRewrite(EncapsulateFieldModel model)
        {
            IEncapsulateFieldRewriteSession refactorRewriteSession = new EncapsulateFieldRewriteSession(RewritingManager.CheckOutCodePaneSession());
            refactorRewriteSession = GeneratePreview(model, refactorRewriteSession);

            var previewRewriter = refactorRewriteSession.CheckOutModuleRewriter(_targetQMN);

            return previewRewriter.GetText(maxConsecutiveNewLines: 3);
        }

        public IEncapsulateFieldRewriteSession GeneratePreview(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession)
        {
            if (!model.SelectedFieldCandidates.Any()) { return refactorRewriteSession; }

            return RefactorRewrite(model, refactorRewriteSession, asPreview: true);
        }

        public IEncapsulateFieldRewriteSession RefactorRewrite(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession)
        {
            if (!model.SelectedFieldCandidates.Any()) { return refactorRewriteSession; }

            return RefactorRewrite(model, refactorRewriteSession, asPreview: false);
        }

        private IEncapsulateFieldRewriteSession RefactorRewrite(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession, bool asPreview)
        {
            ModifyFields(model, refactorRewriteSession);

            ModifyReferences(model, refactorRewriteSession);

            InsertNewContent(model, refactorRewriteSession, asPreview);

            return refactorRewriteSession;
        }

        private void ModifyReferences(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession)
        {
            foreach (var field in model.SelectedFieldCandidates)
            {
                field.ReferenceQualifier = model.ConvertFieldsToUDTMembers
                    ? model.StateUDTField.FieldIdentifier
                    : null;

                field.LoadFieldReferenceContextReplacements();
            }

            foreach (var rewriteReplacement in model.SelectedFieldCandidates.SelectMany(field => field.ReferenceReplacements))
            {
                (ParserRuleContext Context, string Text) = rewriteReplacement.Value;
                var rewriter = refactorRewriteSession.CheckOutModuleRewriter(rewriteReplacement.Key.QualifiedModuleName);
                rewriter.Replace(Context, Text);
            }
        }

        private void ModifyFields(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession)
        {
            if (model.ConvertFieldsToUDTMembers)
            {
                IModuleRewriter rewriter;

                foreach (var field in model.SelectedFieldCandidates)
                {
                    rewriter = refactorRewriteSession.CheckOutModuleRewriter(_targetQMN);

                    refactorRewriteSession.Remove(field.Declaration, rewriter);
                }

                if (!model.StateUDTField.IsExistingDeclaration)
                {
                    return;
                }

                var stateUDT = model.StateUDTField;

                stateUDT.AddMembers(model.SelectedFieldCandidates);

                rewriter = refactorRewriteSession.CheckOutModuleRewriter(_targetQMN);

                rewriter.Replace(stateUDT.AsTypeDeclaration, stateUDT.TypeDeclarationBlock(_indenter));

                return;
            }

            foreach (var field in model.SelectedFieldCandidates)
            {
                var rewriter = refactorRewriteSession.CheckOutModuleRewriter(_targetQMN);

                if (field.Declaration.HasPrivateAccessibility() && field.FieldIdentifier.Equals(field.Declaration.IdentifierName))
                {
                    rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
                    continue;
                }

                if (field.Declaration.IsDeclaredInList() && !field.Declaration.HasPrivateAccessibility())
                {
                    refactorRewriteSession.Remove(field.Declaration, rewriter);
                    continue;
                }

                rewriter.Rename(field.Declaration, field.FieldIdentifier);
                rewriter.SetVariableVisiblity(field.Declaration, Accessibility.Private.TokenString());
                rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
            }
        }

        private void InsertNewContent(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession, bool postPendPreviewMessage = false)
        {
            _newContent = new Dictionary<NewContentTypes, List<string>>
            {
                { NewContentTypes.PostContentMessage, new List<string>() },
                { NewContentTypes.DeclarationBlock, new List<string>() },
                { NewContentTypes.MethodBlock, new List<string>() },
                { NewContentTypes.TypeDeclarationBlock, new List<string>() }
            };

            var rewriter = refactorRewriteSession.CheckOutModuleRewriter(_targetQMN);

            LoadNewDeclarationBlocks(model);

            LoadNewPropertyBlocks(model);

            if (postPendPreviewMessage)
            {
                _newContent[NewContentTypes.PostContentMessage].Add(EncapsulateFieldResources.PreviewEndOfChangesMarker);
            }

            var newContentBlock = string.Join(DoubleSpace,
                            (_newContent[NewContentTypes.TypeDeclarationBlock])
                            .Concat(_newContent[NewContentTypes.DeclarationBlock])
                            .Concat(_newContent[NewContentTypes.MethodBlock])
                            .Concat(_newContent[NewContentTypes.PostContentMessage]))
                        .Trim();


            if (_codeSectionStartIndex.HasValue)
            {
                rewriter.InsertBefore(_codeSectionStartIndex.Value, $"{newContentBlock}{DoubleSpace}");
            }
            else
            {
                rewriter.InsertAtEndOfFile($"{DoubleSpace}{newContentBlock}");
            }
        }

        private void LoadNewDeclarationBlocks(EncapsulateFieldModel model)
        {
            if (model.ConvertFieldsToUDTMembers)
            {
                if (model.StateUDTField?.IsExistingDeclaration ?? false) { return; }

                model.StateUDTField = _encapsulationCandidateFactory.CreateStateUDTField();

                model.StateUDTField.AddMembers(model.SelectedFieldCandidates);

                AddCodeBlock(NewContentTypes.TypeDeclarationBlock, model.StateUDTField.TypeDeclarationBlock(_indenter));
                AddCodeBlock(NewContentTypes.DeclarationBlock, model.StateUDTField.FieldDeclarationBlock);
                return;
            }

            //New field declarations created here were removed from their list within ModifyFields(...)
            var fieldsRequiringNewDeclaration = model.SelectedFieldCandidates
                .Where(field => field.Declaration.IsDeclaredInList()
                                    && field.Declaration.Accessibility != Accessibility.Private);

            foreach (var field in fieldsRequiringNewDeclaration)
            {
                var targetIdentifier = field.Declaration.Context.GetText().Replace(field.IdentifierName, field.FieldIdentifier);
                var newField = field.Declaration.IsTypeSpecified
                    ? $"{Tokens.Private} {targetIdentifier}"
                    : $"{Tokens.Private} {targetIdentifier} {Tokens.As} {field.Declaration.AsTypeName}";

                AddCodeBlock(NewContentTypes.DeclarationBlock, newField);
            }
        }

        private void LoadNewPropertyBlocks(EncapsulateFieldModel model)
        {
            var propertyGenerationSpecs = model.SelectedFieldCandidates
                                                .SelectMany(f => f.PropertyAttributeSets);

            var generator = new PropertyGenerator();
            foreach (var spec in propertyGenerationSpecs)
            {
                AddCodeBlock(NewContentTypes.MethodBlock, generator.AsPropertyBlock(spec, _indenter));
            }
        }

        private void AddCodeBlock(NewContentTypes contentType, string block)
            => _newContent[contentType].Add(block);
    }
}
