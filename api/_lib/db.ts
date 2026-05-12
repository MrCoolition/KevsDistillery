import { neon } from '@neondatabase/serverless';
import type { DossierSummary } from './package-builder.js';
import type { DiscoveryModel } from './types.js';

export async function persistDossierRun(
  model: DiscoveryModel,
  zipBuffer: Buffer,
  summary: DossierSummary,
): Promise<{ persisted: boolean; error?: string }> {
  const databaseUrl = process.env.DATABASE_URL;
  if (!databaseUrl) {
    model.limitations.push('DATABASE_URL is not configured. The package was generated but not persisted to Neon.');
    return { persisted: false };
  }

  try {
    const sql = neon(databaseUrl);
    await sql`set search_path = data_discovery, public`;
    await sql`
      insert into discovery_jobs (
        job_id,
        package_name,
        source_process_name,
        status,
        generated_date,
        source_file_count,
        object_count,
        query_count,
        macro_count,
        linked_source_count,
        lineage_blocker_count,
        qa_status,
        canonical_model
      )
      values (
        ${model.runId},
        ${model.packageName},
        ${model.sourceProcessName},
        'COMPLETED',
        ${model.generatedDate},
        ${summary.fileCount},
        ${summary.objectCount},
        ${summary.queryCount},
        ${summary.macroCount},
        ${summary.linkedSourceCount},
        ${summary.lineageBlockers},
        ${summary.qaStatus},
        ${JSON.stringify(redactEvidenceContent(model))}
      )
      on conflict (job_id) do update set
        status = excluded.status,
        canonical_model = excluded.canonical_model,
        qa_status = excluded.qa_status,
        updated_at = now()
    `;

    for (const source of model.sourceFiles) {
      await sql`
        insert into discovery_source_files (
          source_id,
          job_id,
          file_name,
          file_type,
          extension,
          file_size_bytes,
          sha256,
          evidence_id
        )
        values (
          ${source.source_id},
          ${model.runId},
          ${source.file_name},
          ${source.file_type},
          ${source.extension},
          ${source.file_size_bytes},
          ${source.sha256},
          ${source.evidence_id}
        )
        on conflict (source_id) do nothing
      `;
    }

    for (const evidence of model.evidence) {
      await sql`
        insert into discovery_evidence_items (
          evidence_id,
          job_id,
          title,
          category,
          relative_path,
          source_file,
          confidence,
          summary
        )
        values (
          ${evidence.evidence_id},
          ${model.runId},
          ${evidence.title},
          ${evidence.category},
          ${evidence.relative_path},
          ${evidence.source_file},
          ${evidence.confidence},
          ${evidence.summary}
        )
        on conflict (job_id, evidence_id) do nothing
      `;
    }

    for (const node of model.nodes) {
      await sql`
        insert into lineage_nodes (
          node_id,
          job_id,
          node_type,
          name,
          description,
          source_file,
          business_purpose,
          owner_status,
          criticality,
          confidence,
          evidence_id,
          recommended_action,
          failure_impact,
          dollar_exposure
        )
        values (
          ${node.node_id},
          ${model.runId},
          ${node.node_type},
          ${node.name},
          ${node.description},
          ${node.source_file},
          ${node.business_purpose},
          ${node.owner_status},
          ${node.criticality},
          ${node.confidence},
          ${node.evidence_id},
          ${node.recommended_action},
          ${node.failure_impact},
          ${node.dollar_exposure}
        )
        on conflict (job_id, node_id) do nothing
      `;
    }

    for (const edge of model.edges) {
      await sql`
        insert into lineage_edges (
          edge_id,
          job_id,
          from_node_id,
          to_node_id,
          edge_type,
          description,
          automated_flag,
          transformation_id,
          cadence,
          confidence,
          evidence_id
        )
        values (
          ${edge.edge_id},
          ${model.runId},
          ${edge.from_node_id},
          ${edge.to_node_id},
          ${edge.edge_type},
          ${edge.description},
          ${edge.automated_flag},
          ${edge.transformation_id || null},
          ${edge.cadence || null},
          ${edge.confidence},
          ${edge.evidence_id}
        )
        on conflict (job_id, edge_id) do nothing
      `;
    }

    for (const action of model.actions) {
      await sql`
        insert into action_items (
          action_id,
          job_id,
          title,
          description,
          source_asset,
          owner_role,
          recommended_owner,
          action_type,
          priority,
          severity,
          dependency,
          due_date_or_phase,
          acceptance_criteria,
          evidence_id,
          related_risk,
          expected_business_value,
          status
        )
        values (
          ${action.action_id},
          ${model.runId},
          ${action.title},
          ${action.description},
          ${action.source_asset},
          ${action.owner_role},
          ${action.recommended_owner},
          ${action.action_type},
          ${action.priority},
          ${action.severity},
          ${action.dependency},
          ${action.due_date_or_phase},
          ${action.acceptance_criteria},
          ${action.evidence_id},
          ${action.related_risk},
          ${action.expected_business_value},
          ${action.status}
        )
        on conflict (job_id, action_id) do nothing
      `;
    }

    for (const exposure of model.financialExposure) {
      await sql`
        insert into financial_exposures (
          job_id,
          process_or_output,
          failure_scenario,
          frequency,
          units_affected,
          dollar_per_unit,
          revenue_at_risk,
          margin_percent,
          margin_at_risk,
          rework_hours,
          labor_rate,
          labor_recovery_cost,
          customer_sla_exposure,
          compliance_exposure,
          cash_timing_cost,
          low_impact,
          base_impact,
          high_impact,
          annualized_low,
          annualized_base,
          annualized_high,
          confidence,
          assumptions,
          evidence_id,
          finance_validation_needed
        )
        values (
          ${model.runId},
          ${exposure.process_or_output},
          ${exposure.failure_scenario},
          ${exposure.frequency},
          ${exposure.units_affected},
          ${exposure.dollar_per_unit},
          ${exposure.revenue_at_risk},
          ${exposure.margin_percent},
          ${exposure.margin_at_risk},
          ${exposure.rework_hours},
          ${exposure.labor_rate},
          ${exposure.labor_recovery_cost},
          ${exposure.customer_sla_exposure},
          ${exposure.compliance_exposure},
          ${exposure.cash_timing_cost},
          ${exposure.low_impact},
          ${exposure.base_impact},
          ${exposure.high_impact},
          ${exposure.annualized_low},
          ${exposure.annualized_base},
          ${exposure.annualized_high},
          ${exposure.confidence},
          ${exposure.assumptions},
          ${exposure.evidence_id},
          ${exposure.finance_validation_needed}
        )
      `;
    }

    for (const qa of model.qaRecords) {
      await sql`
        insert into qa_checks (
          qa_id,
          job_id,
          check_text,
          status,
          evidence_id,
          notes
        )
        values (
          ${qa.qa_id},
          ${model.runId},
          ${qa.check},
          ${qa.status},
          ${qa.evidence_id},
          ${qa.notes}
        )
        on conflict (job_id, qa_id) do nothing
      `;
    }

    await sql`
      insert into dossier_packages (
        package_id,
        job_id,
        package_name,
        zip_bytes,
        byte_size,
        qa_status
      )
      values (
        gen_random_uuid(),
        ${model.runId},
        ${model.packageName},
        ${zipBuffer},
        ${zipBuffer.byteLength},
        ${summary.qaStatus}
      )
    `;

    for (const usage of summary.aiCost.models) {
      await sql`
        insert into ai_usage_events (
          job_id,
          model,
          requests,
          input_tokens,
          cached_input_tokens,
          billable_input_tokens,
          output_tokens,
          reasoning_tokens,
          total_tokens,
          input_cost_usd,
          cached_input_cost_usd,
          output_cost_usd,
          total_cost_usd,
          cache_savings_usd,
          cache_hit_rate,
          pricing_source,
          optimization_note
        )
        values (
          ${model.runId},
          ${usage.model},
          ${usage.requests},
          ${usage.inputTokens},
          ${usage.cachedInputTokens},
          ${usage.billableInputTokens},
          ${usage.outputTokens},
          ${usage.reasoningTokens},
          ${usage.totalTokens},
          ${usage.inputCostUsd},
          ${usage.cachedInputCostUsd},
          ${usage.outputCostUsd},
          ${usage.totalCostUsd},
          ${usage.cacheSavingsUsd},
          ${usage.cacheHitRate},
          ${summary.aiCost.pricingSource},
          ${summary.aiCost.optimizationNote}
        )
      `;
    }

    return { persisted: true };
  } catch (error) {
    const message = error instanceof Error ? error.message : 'unknown Neon persistence error';
    model.limitations.push(`Neon persistence failed: ${message}`);
    return { persisted: false, error: message };
  }
}

function redactEvidenceContent(model: DiscoveryModel): DiscoveryModel {
  return {
    ...model,
    evidence: model.evidence.map(({ content, ...evidence }) => ({
      ...evidence,
      content: Buffer.isBuffer(content) ? `[${content.byteLength} bytes]` : `[${String(content).length} characters]`,
    })),
  };
}
